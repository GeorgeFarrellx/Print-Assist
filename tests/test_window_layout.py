from __future__ import annotations

import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from tkinter import ttk
from unittest.mock import patch

import fitz

from print_assist.app import PrintAssistApp
from print_assist.pdf_split_window import PdfSplitWindow
from print_assist.preview_window import PreviewWindow


class WindowLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        try:
            root = tk.Tk()
            root.destroy()
        except tk.TclError as exc:
            raise unittest.SkipTest(f"Tk display is unavailable: {exc}") from exc

    def _buttons(self, parent: tk.Misc) -> list[ttk.Button]:
        buttons = []
        for child in parent.winfo_children():
            if isinstance(child, ttk.Button):
                buttons.append(child)
            buttons.extend(self._buttons(child))
        return buttons

    def _assert_visible(self, window: tk.Misc, buttons: list[ttk.Button]) -> None:
        window.update()
        left, top = window.winfo_rootx(), window.winfo_rooty()
        right, bottom = left + window.winfo_width(), top + window.winfo_height()
        for button in buttons:
            with self.subTest(button=button.cget("text"), size=(window.winfo_width(), window.winfo_height())):
                self.assertTrue(button.winfo_ismapped(), "Button is hidden")
                self.assertEqual(button.winfo_width(), button.winfo_reqwidth(), "Button is clipped horizontally")
                self.assertEqual(button.winfo_height(), button.winfo_reqheight(), "Button is clipped vertically")
                self.assertGreaterEqual(button.winfo_rootx(), left)
                self.assertGreaterEqual(button.winfo_rooty(), top)
                self.assertLessEqual(button.winfo_rootx() + button.winfo_width(), right)
                self.assertLessEqual(button.winfo_rooty() + button.winfo_height(), bottom)

    def _source_pdf(self, folder: Path) -> Path:
        source = folder / "sample.pdf"
        with fitz.open() as document:
            for index in range(4):
                page = document.new_page()
                page.insert_text((40, 40), f"Sample page {index + 1}")
            document.save(source)
        return source

    def test_main_controls_stay_visible_in_short_windows_and_larger_text(self) -> None:
        for scaling in (1.333, 2.0, 2.667):
            root = tk.Tk()
            root.tk.call("tk", "scaling", scaling)
            with patch.object(PrintAssistApp, "_wire_drag_drop", return_value=False), patch("print_assist.app.ole_initialize", return_value=False):
                app = PrintAssistApp(root)
            try:
                app.output_var.set("Output: C:/" + "long-folder-name/" * 40 + "output.pdf")
                for geometry in ("900x720", "900x520", "640x420", "320x200"):
                    root.geometry(geometry)
                    self._assert_visible(root, list(app.buttons.values()))
                    self.assertEqual(app.footer.winfo_rooty() + app.footer.winfo_height(), root.winfo_rooty() + root.winfo_height() - 12)
            finally:
                app._on_close()

    def test_preview_controls_stay_visible_after_resize(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = self._source_pdf(Path(temporary))
            for scaling in (1.333, 2.0, 2.667):
                root = tk.Tk()
                root.tk.call("tk", "scaling", scaling)
                view = PreviewWindow(root, source, Path(temporary) / "output.pdf", lambda _path: None, lambda _status: None, lambda: None)
                try:
                    for geometry in ("1000x760", "800x520", "640x480", "320x200"):
                        root.geometry(geometry)
                        self._assert_visible(root, self._buttons(view.footer))
                finally:
                    view.close()
                    root.destroy()

    def test_splitter_export_and_page_controls_stay_visible_after_resize(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            source = self._source_pdf(Path(temporary))
            for scaling in (1.333, 2.0, 2.667):
                root = tk.Tk()
                root.tk.call("tk", "scaling", scaling)
                view = PdfSplitWindow(root, source, lambda _result: None, lambda: None)
                try:
                    deadline = time.monotonic() + 5
                    while view.info is None and time.monotonic() < deadline:
                        root.update()
                        time.sleep(0.01)
                    self.assertIsNotNone(view.info)
                    view.ranges_var.set("1-2, 3-4")
                    view._update_plan()
                    for geometry in ("1100x800", "950x620", "900x540", "320x200"):
                        view.window.geometry(geometry)
                        self._assert_visible(view.window, self._buttons(view.window))
                finally:
                    view.close()
                    root.destroy()


if __name__ == "__main__":
    unittest.main()
