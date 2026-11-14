"""Regression tests for the six defects fixed after the 0.2.0 review."""

import gc
import os
import threading
import time
import unittest
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch
from xml.etree import ElementTree
from zipfile import ZipFile

from converter import (
    EDITOR_THRESHOLD,
    Cancelled,
    ConversionError,
    DiskRows,
    JsonNumber,
    cancellation,
    exceeds_editor_limit,
    parse_table,
    read_table_file,
)
from exporters import SCRATCH_LOCK, ExportPermissionError, save_table, save_workbook
from test_converter import scratch_folder

ODF = "urn:oasis:names:tc:opendocument:xmlns:"
TABLE = f"{{{ODF}table:1.0}}"


def _disk_table(rows: int = 40):
    import json

    source = "\n".join(json.dumps({"id": index, "name": f"이름{index}"}) for index in range(rows)) + "\n"
    with patch("converter.EDITOR_THRESHOLD", 0):
        table = parse_table(source, "jsonl")
    assert isinstance(table.rows, DiskRows)
    return table


class EditorThresholdTests(TestCase):
    def test_threshold_counts_utf8_bytes_not_characters(self):
        korean = "가" * (EDITOR_THRESHOLD // 3 + 1)  # ~3.5 M characters, just over 10 MiB in UTF-8
        self.assertLess(len(korean), EDITOR_THRESHOLD)
        self.assertTrue(exceeds_editor_limit(korean))
        self.assertFalse(exceeds_editor_limit("a" * EDITOR_THRESHOLD))
        self.assertTrue(exceeds_editor_limit("a" * (EDITOR_THRESHOLD + 1)))
        self.assertTrue(exceeds_editor_limit("😀" * 3, limit=11))
        self.assertFalse(exceeds_editor_limit("😀" * 3, limit=12))
        self.assertFalse(exceeds_editor_limit("\ud800", limit=3))  # lone surrogates do not raise

    def test_pasted_multibyte_text_over_limit_is_spooled_to_disk(self):
        record = '{"t":"' + "가" * 40 + '"}'
        self.assertLess(len(record), 150)
        with patch("converter.exceeds_editor_limit", lambda source: len(source.encode()) > 100):
            table = parse_table(record)
        self.assertIsInstance(table.rows, DiskRows)
        self.assertEqual(list(table.rows), [{0: "가" * 40}])
        table.close()


class DiskRowsTests(TestCase):
    def test_interleaved_and_nested_iterations_keep_their_own_position(self):
        table = _disk_table()
        rows = table.rows
        first, second = iter(rows), iter(rows)
        self.assertEqual(next(first)[0], 0.0)
        self.assertEqual(next(first)[0], 1.0)
        self.assertEqual(next(second)[0], 0.0)
        self.assertEqual(next(first)[0], 2.0)
        self.assertEqual([row[0] for row in second], [float(n) for n in range(1, 40)])
        self.assertEqual([row[0] for row in first], [float(n) for n in range(3, 40)])
        nested = [(outer[0], sum(1 for _ in rows)) for outer in rows]
        self.assertEqual(nested, [(float(n), 40) for n in range(40)])
        table.close()

    def test_parallel_threads_read_complete_rows(self):
        table = _disk_table(2000)
        results, errors = [], []

        def read():
            try:
                results.append([row[1] for row in table.rows])
            except Exception as exc:  # pragma: no cover - reported below
                errors.append(exc)

        threads = [threading.Thread(target=read) for _ in range(6)]
        with patch("converter.READ_CHUNK", 97):
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()
        self.assertEqual(errors, [])
        expected = [f"이름{n}" for n in range(2000)]
        self.assertEqual(results, [expected] * 6)
        table.close()

    def test_explicit_close_context_manager_and_collection(self):
        table = _disk_table()
        file = table.rows.file
        with table.rows as rows:
            self.assertFalse(rows.closed)
        self.assertTrue(rows.closed and file.closed)
        table.close()  # idempotent
        with self.assertRaises(ValueError):
            next(iter(rows))
        collected = _disk_table()
        file = collected.rows.file
        del collected
        gc.collect()
        self.assertTrue(file.closed)


class NumberTests(TestCase):
    def test_equal_values_get_the_same_cell_type(self):
        for literal in ("1e16", "1E+16", "1.0e16", "10000000000000000", "10000000000000000.000"):
            with self.subTest(literal=literal):
                self.assertEqual(JsonNumber(literal).to_cell(), 1e16)
        for literal in ("123456789012345e3", "123456789012345000"):
            self.assertEqual(JsonNumber(literal).to_cell(), 1.23456789012345e17)
        self.assertEqual(JsonNumber("0.10000000000000000000").to_cell(), 0.1)
        self.assertEqual(JsonNumber("-0").to_cell(), 0.0)

    def test_values_needing_more_than_15_digits_stay_text(self):
        for literal in (
            "9007199254740993", "1234567890123456e10", "0.1234567890123456",
            "1000000000000000000000000000001",  # would round to 1e30 under Decimal.normalize()
            "1e400", "1e-400",
        ):
            with self.subTest(literal=literal):
                self.assertEqual(JsonNumber(literal).to_cell(), literal)
        self.assertEqual(
            parse_table('[{"a":1e16},{"a":10000000000000000}]').rows,
            [{0: 1e16}, {0: 1e16}],
        )


class SaveErrorTests(TestCase):
    def test_locked_destination_becomes_a_conversion_error_with_os_details(self):
        with scratch_folder() as folder:
            destination = folder / "report.xlsx"
            destination.write_bytes(b"previous version")
            locked = PermissionError(13, "Permission denied", str(destination))
            with patch("exporters.os.replace", side_effect=locked) as replace, patch("exporters.REPLACE_DELAY", 0):
                with self.assertRaises(ConversionError) as raised:
                    save_table(parse_table('{"a":1}'), destination)
            error = raised.exception
            self.assertIsInstance(error, PermissionError)
            self.assertIsInstance(error, ExportPermissionError)
            self.assertEqual((error.code, error.details), ("locked", {"name": "report.xlsx"}))
            self.assertEqual((error.errno, error.filename), (13, str(destination)))
            self.assertIs(error.__cause__, locked)
            self.assertGreater(replace.call_count, 1)
            self.assertEqual(destination.read_bytes(), b"previous version")
            self.assertEqual(list(folder.iterdir()), [destination])

    def test_transient_lock_is_retried(self):
        with scratch_folder() as folder:
            destination = folder / "report.csv"
            calls = []
            real_replace = os.replace

            def flaky(source, target):
                calls.append(source)
                if len(calls) < 3:
                    raise PermissionError(13, "Permission denied")
                real_replace(source, target)

            with patch("exporters.os.replace", flaky), patch("exporters.REPLACE_DELAY", 0):
                save_table(parse_table('{"a":1}'), destination)
            self.assertEqual(len(calls), 3)
            self.assertTrue(destination.read_bytes().startswith(b"\xef\xbb\xbf"))
            self.assertEqual(list(folder.iterdir()), [destination])

    @unittest.skipUnless(os.name == "nt", "Windows share modes")
    def test_file_open_without_sharing_like_excel(self):
        import ctypes
        from ctypes import wintypes

        create = ctypes.windll.kernel32.CreateFileW
        create.restype = wintypes.HANDLE
        close = ctypes.windll.kernel32.CloseHandle
        with scratch_folder() as folder:
            destination = folder / "open-in-excel.xlsx"
            destination.write_bytes(b"previous version")
            handle = create(str(destination), 0x80000000, 0, None, 3, 0x80, None)  # GENERIC_READ, no sharing
            self.assertNotEqual(handle, wintypes.HANDLE(-1).value)
            try:
                with patch("exporters.REPLACE_DELAY", 0), self.assertRaises(ExportPermissionError) as raised:
                    save_table(parse_table('{"a":1}'), destination)
            finally:
                close(handle)
            self.assertEqual(raised.exception.code, "locked")
            self.assertEqual(destination.read_bytes(), b"previous version")
            self.assertEqual(list(folder.iterdir()), [destination])

    def test_other_os_errors_are_reported_as_save_errors(self):
        with scratch_folder() as folder:
            destination = folder / "full.ods"
            with patch("exporters._write_ods", side_effect=OSError(28, "No space left on device")):
                with self.assertRaises(ConversionError) as raised:
                    save_table(parse_table('{"a":1}'), destination)
            self.assertEqual(raised.exception.code, "save")
            self.assertIsInstance(raised.exception, OSError)
            self.assertEqual(raised.exception.errno, 28)
            self.assertEqual(list(folder.iterdir()), [])


class CancellationTests(TestCase):
    def test_cancelled_save_keeps_destination_and_leaves_no_files(self):
        table = _disk_table(500)
        for extension in (".xlsx", ".xltm", ".ods", ".csv"):
            with self.subTest(extension=extension), scratch_folder() as folder:
                destination = folder / ("sample" + extension)
                destination.write_bytes(b"previous version")
                event = threading.Event()
                event.set()
                with cancellation(event), self.assertRaises(Cancelled):
                    save_table(table, destination)
                self.assertEqual(destination.read_bytes(), b"previous version")
                self.assertEqual(list(folder.iterdir()), [destination])
        table.close()

    def test_cancelled_large_read_stops(self):
        with scratch_folder() as folder:
            path = folder / "big.jsonl"
            path.write_text('{"a":1}\n' * 1000, encoding="utf-8")
            event = threading.Event()
            event.set()
            with cancellation(event), self.assertRaises(Cancelled):
                read_table_file(path, editor_limit=0)
            _, table = read_table_file(path, editor_limit=0)  # outside the block nothing is cancelled
            self.assertEqual(len(table.rows), 1000)
            table.close()

    def test_scratch_left_by_a_killed_export_is_removed_by_the_next_one(self):
        with scratch_folder() as folder:
            orphan = folder / ".tablyfern-dead.partial"
            orphan.mkdir()
            (orphan / SCRATCH_LOCK).write_bytes(b"")
            (orphan / "output.xlsx").write_bytes(b"x" * 100)
            (orphan / "tmpabc").write_bytes(b"xlsxwriter scratch")
            unrelated = folder / "keep.partial"
            unrelated.mkdir()
            destination = folder / "out.ods"
            if os.name != "nt":
                old = time.time() - 7 * 3600
                os.utime(orphan, (old, old))
            save_workbook({"A": parse_table('{"a":1}')}, destination)
            self.assertEqual(sorted(p.name for p in folder.iterdir()), ["keep.partial", "out.ods"])
            unrelated.rmdir()

    @unittest.skipUnless(os.name == "nt", "open files cannot be deleted on Windows")
    def test_scratch_of_a_running_export_is_not_removed(self):
        with scratch_folder() as folder:
            live = folder / ".tablyfern-live.partial"
            live.mkdir()
            with (live / SCRATCH_LOCK).open("wb"):
                save_table(parse_table('{"a":1}'), folder / "out.csv")
                self.assertTrue(live.is_dir())
            (live / SCRATCH_LOCK).unlink()
            live.rmdir()


class OdsTests(TestCase):
    def test_rows_are_padded_to_the_full_width(self):
        table = parse_table('[{"a":1,"b":2,"c":3,"d":4},{"a":1},{"b":2},{},{"d":4}]')
        with scratch_folder() as folder:
            destination = folder / "sparse.ods"
            save_table(table, destination)
            with ZipFile(destination) as archive:
                content = ElementTree.fromstring(archive.read("content.xml"))
        sheet = content.find(f".//{TABLE}table")
        columns = sheet.findall(f"{TABLE}table-column")
        self.assertEqual(sum(int(c.get(f"{TABLE}number-columns-repeated", "1")) for c in columns), 4)
        self.assertEqual(list(sheet)[0].tag, f"{TABLE}table-column")
        widths = [
            sum(int(cell.get(f"{TABLE}number-columns-repeated", "1")) for cell in row)
            for row in sheet.findall(f"{TABLE}table-row")
        ]
        self.assertEqual(widths, [4] * 6)


@unittest.skipIf(os.environ.get("TABLYFERN_SKIP_GUI"), "GUI tests disabled")
class WindowTests(TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication

        cls.app = QApplication.instance() or QApplication([])

    def wait(self, condition, timeout=10.0):
        deadline = time.monotonic() + timeout
        while not condition():
            self.app.processEvents()
            if time.monotonic() > deadline:
                self.fail("timed out")
            time.sleep(0.01)
        self.app.processEvents()

    def test_large_multibyte_paste_is_parsed_off_the_gui_thread(self):
        from app import Window

        window = Window()
        text = '"' + "가" * (EDITOR_THRESHOLD // 3 + 10) + '"'
        self.assertLess(len(text), EDITOR_THRESHOLD)
        window.editor.setPlainText(text)
        with patch.object(Window, "start_job") as start_job, patch("app.parse_table") as parse:
            window.update_preview()
        start_job.assert_called_once()
        parse.assert_not_called()
        window.deleteLater()

    def test_cancel_button_stops_a_job(self):
        from app import Window

        window = Window()
        started = threading.Event()

        def action():
            from converter import checkpoint

            started.set()
            while True:
                checkpoint()
                time.sleep(0.01)

        window.start_job("loading", action, lambda _: self.fail("must not succeed"), "read")
        self.assertTrue(window.cancel_button.isVisibleTo(window))
        self.assertFalse(window.content.isEnabled())
        self.assertTrue(window.cancel_button.isEnabled())
        started.wait(5)
        window.cancel_button.click()
        self.wait(lambda: window.job is None)
        self.assertEqual(window.status_key, "cancelled")
        self.assertTrue(window.content.isEnabled())
        self.assertFalse(window.cancel_button.isVisibleTo(window))
        window.deleteLater()

    def test_close_during_job_asks_then_cancels_and_closes(self):
        from app import Window

        window = Window()
        window.show()
        table = _disk_table()
        window.show_table(table)

        def action():
            from converter import checkpoint

            while True:
                checkpoint()
                time.sleep(0.01)

        window.start_job("saving", action, lambda _: None, "save")
        with patch.object(Window, "confirm_quit", return_value=False) as asked:
            window.close()
        asked.assert_called_once()
        self.assertTrue(window.isVisible())
        self.assertIsNotNone(window.job)
        with patch.object(Window, "confirm_quit", return_value=True):
            window.close()
        self.assertTrue(window.quit_requested)
        self.assertEqual(window.status_key, "cancelling")
        self.wait(lambda: not window.isVisible())
        self.assertIsNone(window.job)
        self.assertTrue(table.rows.closed)
        window.deleteLater()

    def test_replaced_tables_release_their_files(self):
        from app import Window

        window = Window()
        first, second = _disk_table(), _disk_table()
        window.show_table(first)
        window.show_table(second)
        self.assertTrue(first.rows.closed)
        self.assertFalse(second.rows.closed)
        window.load_example()
        self.assertTrue(second.rows.closed)
        window.deleteLater()


if __name__ == "__main__":
    unittest.main()
