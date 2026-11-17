"""Exercise preview and export state through the desktop controls."""

import time
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch
from zipfile import ZipFile

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QFileDialog

from app import Window
from test_converter import scratch_folder


class WorkspaceTests(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.window = Window()
        self.window.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen)
        self.window.show()
        self.app.processEvents()

    def tearDown(self):
        self.wait_for_job()
        self.window.close()
        self.window.deleteLater()

    def wait_for_job(self):
        deadline = time.monotonic() + 10
        while self.window.job is not None and time.monotonic() < deadline:
            QTest.qWait(10)
        self.assertIsNone(self.window.job)

    def test_edited_input_must_be_previewed_before_it_can_be_saved(self):
        window = self.window
        preview = window.buttons["preview"]
        save = window.buttons["export"]
        self.assertTrue(preview.isHidden())
        self.assertFalse(save.isEnabled())

        window.buttons["sample"].click()
        self.assertTrue(preview.isHidden())
        self.assertTrue(save.isEnabled())
        self.assertEqual(window.grid.item(0, 0).text(), "Mina")

        window.editor.setPlainText('{"item":')
        self.assertTrue(preview.isVisible())
        self.assertTrue(window.shortcuts["preview"].isEnabled())
        self.assertFalse(save.isEnabled())
        self.assertEqual(window.grid.rowCount(), 0)
        preview.click()
        self.assertTrue(window.status_error)
        self.assertIn("Invalid JSON", window.status.text())
        self.assertTrue(preview.isVisible())
        self.assertFalse(window.shortcuts["export"].isEnabled())

        window.language.setCurrentIndex(1)
        self.assertIn("오류", window.status.text())
        window.editor.setPlainText('{"item":"verified"}')
        with patch("converter.EDITOR_THRESHOLD", 1):
            preview.click()
            self.assertTrue(window.busy.isVisible())
            self.assertFalse(save.isEnabled())
            self.wait_for_job()
        self.assertTrue(preview.isHidden())
        self.assertFalse(window.shortcuts["preview"].isEnabled())
        self.assertTrue(save.isEnabled())
        with scratch_folder() as folder:
            destination = folder / "output.xlsx"
            with patch.object(QFileDialog, "exec", return_value=True):
                with patch.object(QFileDialog, "selectedFiles", return_value=[str(destination)]):
                    save.click()
                    self.assertTrue(window.busy.isVisible())
                    self.assertFalse(save.isEnabled())
                    self.wait_for_job()
            with ZipFile(destination) as workbook:
                self.assertIn(b"verified", workbook.read("xl/worksheets/sheet1.xml"))
            self.assertIn("저장 완료", window.status.text())
            self.assertTrue(save.isEnabled())

        window.editor.clear()
        self.assertTrue(preview.isHidden())
        self.assertFalse(save.isEnabled())

    def test_opening_files_and_switching_tabs_updates_the_preview_automatically(self):
        window = self.window
        preview = window.buttons["preview"]
        with scratch_folder() as folder:
            inputs = {
                "orders.json": '[{"city":"Boston"}]',
                "events.jsonl": '{"type":"open"}\n{"type":"close"}',
                "bad.json": '{"broken":',
            }
            for name, contents in inputs.items():
                (folder / name).write_text(contents, encoding="utf-8")

            with patch.object(QFileDialog, "getOpenFileName", return_value=(str(folder / "orders.json"), "")):
                window.buttons["open"].click()
                self.wait_for_job()
            self.assertTrue(preview.isHidden())
            self.assertEqual(window.grid.item(0, 0).text(), "Boston")

            paths = [str(folder / name) for name in ("orders.json", "events.jsonl")]
            with patch.object(QFileDialog, "getOpenFileNames", return_value=(paths, "")):
                window.buttons["open_batch"].click()
                self.wait_for_job()
            self.assertTrue(window.sheet_selector.isVisible())
            self.assertTrue(window.editor.isReadOnly())
            self.assertEqual(window.batch_picker.itemText(1), "events")
            window.batch_picker.setCurrentIndex(1)
            self.assertEqual(window.grid.rowCount(), 2)
            self.assertEqual(window.grid.item(1, 0).text(), "close")
            self.assertNotIn("CSV", [window.output_format.itemText(i) for i in range(window.output_format.count())])
            self.assertTrue(preview.isHidden())

            with patch.object(QFileDialog, "getOpenFileName", return_value=(str(folder / "bad.json"), "")):
                window.buttons["open"].click()
                self.wait_for_job()
            self.assertTrue(window.status_error)
            self.assertEqual(window.grid.item(1, 0).text(), "close")
            self.assertTrue(window.buttons["export"].isEnabled())
