from contextlib import contextmanager
from pathlib import Path
from tempfile import gettempdir
from unittest import TestCase
from unittest.mock import patch
from uuid import uuid4
from xml.etree import ElementTree
from zipfile import ZipFile
from xlsxwriter.exceptions import FileSizeError

from converter import BatchInputError, ConversionError, DiskRows, parse_table, read_batch, read_json_file, read_table_file
from exporters import save_table, save_workbook

NAMESPACES = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@contextmanager
def scratch_folder():
    folder = Path(gettempdir()) / f"tablyfern-{uuid4().hex}"
    folder.mkdir()
    try:
        yield folder
    finally:
        for file in folder.iterdir():
            file.unlink()
        folder.rmdir()


class ConverterTests(TestCase):
    def test_package_size_error_keeps_existing_destination(self):
        with scratch_folder() as folder:
            destination = folder / "existing.xlsx"
            destination.write_bytes(b"previous version")
            with patch("exporters._write_xlsx", side_effect=FileSizeError("too large")):
                with self.assertRaises(ConversionError) as rejected:
                    save_table(parse_table('{"id":1}'), destination)
            self.assertEqual(rejected.exception.code, "workbook_size")
            self.assertEqual(destination.read_bytes(), b"previous version")
            self.assertEqual(list(folder.iterdir()), [destination])

    def test_batch_names_keep_order_and_resolve_conflicts(self):
        with scratch_folder() as folder:
            filenames = (
                "logs.json", "LOGS.jsonl", "History.json", "'[서울]'.json",
                "😀" * 17 + "a.json", "😀" * 17 + "b.json", "a&b.json",
            )
            paths = [folder / name for name in filenames]
            for path in paths:
                path.write_text('{"name":"sample"}', encoding="utf-8")
            batch = read_batch(paths)
            self.assertEqual([item.path for item in batch], paths)
            self.assertEqual(
                [item.name for item in batch],
                ["logs", "LOGS (2)", "History (2)", "_서울_", "😀" * 15, "😀" * 13 + " (2)", "a&b"],
            )

    def test_batch_reports_bad_file_and_enforces_limits(self):
        with scratch_folder() as folder:
            good, bad = folder / "good.json", folder / "bad.jsonl"
            good.write_text('{"a":1}', encoding="utf-8-sig")
            bad.write_text('{"a":1}\n{"b":}', encoding="utf-8")
            with self.assertRaises(BatchInputError) as raised:
                read_batch([good, bad])
            self.assertEqual(raised.exception.path, bad)
            self.assertEqual(raised.exception.issue.details, {"line": 2, "column": 6})
            with patch("converter.MAX_BATCH_FILES", 1), self.assertRaises(ConversionError) as count:
                read_batch([good, bad])
            self.assertEqual(count.exception.code, "batch_count")

    def test_large_jsonl_stays_on_disk_and_reassembles_long_cells(self):
        import json

        text = "a" * 32766 + "😀" + "line\n" * 255
        with scratch_folder() as folder:
            path = folder / "large.jsonl"
            with path.open("w", encoding="utf-8") as file:
                for number in range(340):
                    file.write(json.dumps({"id": number, "text": text, "text [part 2]": "literal"}) + "\n")
            self.assertGreater(path.stat().st_size, 10 * 1024 * 1024)
            source, table = read_table_file(path)
            self.assertIsNone(source)
            self.assertIsInstance(table.rows, DiskRows)
            self.assertEqual(len(table.rows), 340)
            self.assertEqual(table.headers, ("id", "text", "text [part 2]_", "text [part 3]", "text [part 2]"))
            for index, row in enumerate(table.rows):
                self.assertEqual(row[0], index)
                self.assertEqual("".join(row[column] for column in (1, 2, 3)), text)
                self.assertEqual(row[4], "literal")
            self.assertEqual(next(iter(table.rows))[0], 0)

    def test_streaming_json_matches_mixed_shapes_and_numeric_chunk_boundaries(self):
        with scratch_folder() as folder:
            path = folder / "array.json"
            source = '[{"id":1}, null, -123456789.125e-3, ["x",2], {"id":9007199254740993}]'
            path.write_text(source, encoding="utf-8-sig")
            with patch("converter.READ_CHUNK", 3):
                editor, table = read_table_file(path, editor_limit=0)
                self.assertIsNone(editor)
                expected = parse_table(source)
                self.assertEqual(table.headers, expected.headers)
                self.assertEqual(list(table.rows), expected.rows)
                for invalid in ("[1,]", "[1", "[1] 2", "[01]"):
                    path.write_text(invalid, encoding="utf-8")
                    with self.subTest(invalid=invalid), self.assertRaises(ConversionError):
                        read_table_file(path, editor_limit=0)

    def test_second_tab_export_failure_preserves_existing_workbook(self):
        with scratch_folder() as folder:
            destination = folder / "sample.ods"
            destination.write_bytes(b"previous version")
            tables = {"Good": parse_table("{}"), "Bad": parse_table(r'{"control":"\u0000"}')}
            with self.assertRaises(ConversionError) as rejected:
                save_workbook(tables, destination)
            self.assertEqual(rejected.exception.details, {"sheet": "Bad", "row": 2, "column": 1})
            self.assertEqual(destination.read_bytes(), b"previous version")
            with self.assertRaises(ConversionError) as delimited:
                save_workbook(tables, folder / "sample.csv")
            self.assertEqual(delimited.exception.code, "batch_format")
            self.assertEqual(list(folder.iterdir()), [destination])

    def test_flattening_keeps_literal_keys_and_nested_arrays(self):
        table = parse_table(r'[{"a.b":1,"a":{"b":2},"tags":[9007199254740993],"x":null},{"late":true}]')
        self.assertEqual(table.headers, (r"a\.b", "a.b", "tags", "x", "late"))
        self.assertEqual(table.rows, [{0: 1.0, 1: 2.0, 2: "[9007199254740993]"}, {4: True}])

    def test_mixed_values_and_empty_records(self):
        self.assertEqual(parse_table('[null, {"id":1}, [2,3]]').rows, [{}, {0: '{"id":1}'}, {0: "[2,3]"}])
        self.assertEqual(parse_table("42").rows, [{0: 42.0}])
        self.assertEqual(parse_table("[{},{}]").rows, [{}, {}])
        self.assertEqual(parse_table("[]").rows, [])

    def test_jsonl_and_auto_detection(self):
        source = '\ufeff{"a":1}\r\n{"b":2}\r\n'
        for mode in ("jsonl", "auto"):
            self.assertEqual(parse_table(source, mode).rows, [{0: 1.0}, {1: 2.0}])
        self.assertEqual(parse_table("[1,2]\n", "jsonl").rows, [{0: "[1,2]"}])
        self.assertEqual(parse_table("[1,2]\n", "auto").rows, [{0: 1.0}, {0: 2.0}])
        with self.assertRaises(ConversionError) as rejected:
            parse_table('{"a":1}\n{"b":}', "jsonl")
        self.assertEqual(rejected.exception.details, {"line": 2, "column": 6})

    def test_incompatible_output_does_not_replace_existing_file(self):
        with scratch_folder() as folder:
            destination = folder / "sample.ods"
            destination.write_bytes(b"previous version")
            with self.assertRaises(ConversionError) as rejected:
                save_table(parse_table(r'{"control":"\u0000"}'), destination)
            self.assertEqual(rejected.exception.code, "ods_character")
            self.assertEqual(destination.read_bytes(), b"previous version")
            self.assertEqual(list(folder.iterdir()), [destination])
            with self.assertRaises(ConversionError) as unsupported:
                save_table(parse_table('{"a":1}'), folder / "sample.xlsb")
            self.assertEqual(unsupported.exception.code, "extension")

    def test_invalid_input_and_cell_limit(self):
        cases = {
            "[": "invalid",
            '{"x":1,"x":2}': "duplicate",
            "[NaN]": "constant",
            r'"\ud800"': "unicode",
            '"' + "😀" * 16_384 + '"': "cell_limit",
        }
        for source, code in cases.items():
            with self.subTest(code=code), self.assertRaises(ConversionError) as raised:
                parse_table(source)
            self.assertEqual(raised.exception.code, code)

    def test_workbook_types_and_blank_rows(self):
        with scratch_folder() as root:
            source = root / "sample.json"
            destination = root / "sample.xlsx"
            source.write_text(
                '[{"한글":"서울","formula":"=2+2","id":9007199254740993,"num":1.25,"bool":false},{}]',
                encoding="utf-8-sig",
            )
            save_table(parse_table(read_json_file(source)), destination)
            with ZipFile(destination) as archive:
                self.assertIsNone(archive.testzip())
                xml = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))

        cells = {cell.attrib["r"]: cell for cell in xml.findall(".//s:c", NAMESPACES)}
        for address, text in (("A2", "서울"), ("B2", "=2+2"), ("C2", "9007199254740993")):
            self.assertEqual(cells[address].attrib["t"], "inlineStr")
            self.assertEqual(cells[address].findtext("s:is/s:t", namespaces=NAMESPACES), text)
        self.assertEqual(cells["D2"].findtext("s:v", namespaces=NAMESPACES), "1.25")
        self.assertEqual(cells["E2"].attrib["t"], "b")
        self.assertIn("A3", cells)
        self.assertIsNone(xml.find(".//s:f", NAMESPACES))

    def test_failed_replace_keeps_destination_and_removes_temporary_file(self):
        with scratch_folder() as folder:
            destination = folder / "sample.xlsx"
            destination.write_bytes(b"previous version")
            with patch("exporters.os.replace", side_effect=PermissionError), self.assertRaises(PermissionError):
                save_table(parse_table('{"name":"sample"}'), destination)
            self.assertEqual(destination.read_bytes(), b"previous version")
            self.assertEqual(list(destination.parent.iterdir()), [destination])
