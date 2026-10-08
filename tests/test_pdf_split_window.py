from __future__ import annotations

import tempfile
import threading
import time
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from print_assist.pdf_split_window import PdfSplitWindow
from print_assist.pdf_splitter import PdfPart, export_split_pdf


class PdfSplitWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        try:
            cls.root = tk.Tk()
            cls.root.withdraw()
        except tk.TclError as exc:
            raise unittest.SkipTest(f"Tk display is unavailable: {exc}") from exc

    @classmethod
    def tearDownClass(cls) -> None:
        cls.root.destroy()

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.source = Path(self.temp_dir.name) / "source.pdf"
        with fitz.open() as document:
            for index in range(6):
                page = document.new_page()
                page.insert_text((40, 40), f"Page {index + 1}")
            document.save(self.source)
        self.results = []
        self.closed = []
        self.view = PdfSplitWindow(self.root, self.source, self.results.append, lambda: self.closed.append(True))
        self.view.window.withdraw()
        self._pump_until(lambda: self.view.info is not None)

    def tearDown(self) -> None:
        if self.view._exporting:
            self._pump_until(lambda: not self.view._exporting)
        self.view.close()
        self.root.update()
        self.temp_dir.cleanup()

    def _pump_until(self, condition) -> None:
        deadline = time.monotonic() + 5
        while not condition() and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertTrue(condition(), "The desktop worker did not finish in time")

    def test_opening_pdf_does_not_choose_any_breaks(self) -> None:
        self.assertEqual(self.view.parts, (PdfPart(1, 6),))
        self.assertEqual(self.view.ranges_var.get(), "1-6")
        self.assertIn("disabled", self.view.export_button.state())

    def test_preview_marking_and_removing_breaks_updates_typed_ranges(self) -> None:
        self.view._show_page(2)
        self.view._add_boundary()
        self.assertEqual(self.view.ranges_var.get(), "1-2, 3-6")
        self.view._show_page(4)
        self.view._add_boundary()
        self.assertEqual(self.view.ranges_var.get(), "1-2, 3-4, 5-6")
        self.view._remove_boundary()
        self.assertEqual(self.view.ranges_var.get(), "1-2, 3-6")
        self.assertEqual(self.view.page_index, 4)

    def test_typed_ranges_update_file_list_and_preview_navigation(self) -> None:
        self.view.ranges_var.set("1, 2-3, 4-6")
        self.assertTrue(self.view._update_plan())
        self.assertEqual(self.view.parts, (PdfPart(1, 1), PdfPart(2, 3), PdfPart(4, 6)))
        self.view.tree.selection_set("1")
        self.view._select_part()
        self.assertEqual(self.view.page_index, 1)
        self.assertIn("File 2, pages 2-3", self.view.page_label_var.get())

    def test_invalid_ranges_disable_export_and_cannot_leave_stale_parts(self) -> None:
        self.view.ranges_var.set("1-3, 4-6")
        self.assertTrue(self.view._update_plan())
        self.view.ranges_var.set("1-2, 4-6")
        self.assertFalse(self.view._update_plan())
        self.assertEqual(self.view.parts, ())
        self.assertIn("disabled", self.view.export_button.state())
        self.view._export()
        self.assertEqual(list(self.source.parent.iterdir()), [self.source])

    def test_background_export_finishes_before_window_can_close(self) -> None:
        self.view.ranges_var.set("1-3, 4-6")
        entered, release = threading.Event(), threading.Event()

        def delayed_export(*args, **kwargs):
            entered.set()
            if not release.wait(5):
                raise RuntimeError("Test export was not released")
            return export_split_pdf(*args, **kwargs)

        with patch("print_assist.pdf_split_window.export_split_pdf", side_effect=delayed_export):
            self.view._export()
            try:
                self.assertTrue(entered.wait(2))
                self.assertFalse(self.view.close())
                self.assertEqual(self.closed, [])
            finally:
                release.set()
            self._pump_until(lambda: not self.view._exporting)
        self.assertEqual(len(self.results), 1)
        self.assertEqual(len(self.results[0].files), 2)
        self.assertTrue(self.view.close())
        self.assertEqual(self.closed, [True])


if __name__ == "__main__":
    unittest.main()
