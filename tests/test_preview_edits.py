from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import fitz

from print_assist.preview_window import PreviewWindow, _update_manifest_after_page_removal


class PreviewEditTests(unittest.TestCase):
    def test_email_trim_shortens_target_and_shifts_later_sources(self) -> None:
        manifest = [
            {
                "source_name": "email.msg",
                "source_extension": ".msg",
                "output_start_page": 1,
                "output_end_page": 4,
                "output_page_count": 4,
            },
            {
                "source_name": "attachment.pdf",
                "source_extension": ".pdf",
                "output_start_page": 5,
                "output_end_page": 7,
                "output_page_count": 3,
            },
        ]

        updated = _update_manifest_after_page_removal(
            manifest,
            target_entry_index=0,
            kept_end_page=2,
            removed_count=2,
        )

        self.assertEqual(updated[0]["output_end_page"], 2)
        self.assertEqual(updated[0]["output_page_count"], 2)
        self.assertEqual(updated[1]["output_start_page"], 3)
        self.assertEqual(updated[1]["output_end_page"], 5)
        self.assertEqual(manifest[0]["output_end_page"], 4)

    @staticmethod
    def _editor_for_pdf(
        pdf_path: Path,
        manifest: list[dict[str, object]],
        page_index: int,
    ) -> PreviewWindow:
        editor = PreviewWindow.__new__(PreviewWindow)
        editor.preview_pdf_path = pdf_path
        editor.doc = fitz.open(pdf_path)
        editor.file_manifest = manifest
        editor.page_index = page_index
        editor.saved_final = False
        editor._undo_stack = []
        editor.on_status_change = lambda status: None
        editor._render_page = lambda: None
        return editor

    def test_image_crop_keeps_the_complete_document_page_order(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / "preview.pdf"
            source = fitz.open()
            for label in ("before", "image", "after"):
                page = source.new_page(width=300, height=400)
                page.insert_text((40, 80), label)
            source.save(pdf_path)
            source.close()

            manifest = [
                {
                    "source_name": "before.pdf",
                    "source_extension": ".pdf",
                    "output_start_page": 1,
                    "output_end_page": 1,
                    "output_page_count": 1,
                },
                {
                    "source_name": "photo.jpg",
                    "source_extension": ".jpg",
                    "output_start_page": 2,
                    "output_end_page": 2,
                    "output_page_count": 1,
                },
                {
                    "source_name": "after.pdf",
                    "source_extension": ".pdf",
                    "output_start_page": 3,
                    "output_end_page": 3,
                    "output_page_count": 1,
                },
            ]
            editor = self._editor_for_pdf(pdf_path, manifest, page_index=1)
            try:
                editor._apply_image_crop(fitz.Rect(20, 20, 220, 260))
                self.assertEqual(editor.doc.page_count, 3)
                self.assertIn("before", editor.doc[0].get_text())
                self.assertIn("after", editor.doc[2].get_text())
                self.assertEqual(editor.file_manifest, manifest)
            finally:
                editor.doc.close()

    def test_email_trim_removes_only_later_email_pages(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / "preview.pdf"
            source = fitz.open()
            for label in ("email 1", "email 2", "email 3", "attachment"):
                page = source.new_page(width=300, height=400)
                page.insert_text((40, 80), label)
            source.save(pdf_path)
            source.close()

            manifest = [
                {
                    "source_name": "email.msg",
                    "source_extension": ".msg",
                    "output_start_page": 1,
                    "output_end_page": 3,
                    "output_page_count": 3,
                },
                {
                    "source_name": "attachment.pdf",
                    "source_extension": ".pdf",
                    "output_start_page": 4,
                    "output_end_page": 4,
                    "output_page_count": 1,
                },
            ]
            editor = self._editor_for_pdf(pdf_path, manifest, page_index=1)
            try:
                editor._apply_email_trim(180)
                self.assertEqual(editor.doc.page_count, 3)
                self.assertIn("email 1", editor.doc[0].get_text())
                self.assertIn("email 2", editor.doc[1].get_text())
                self.assertIn("attachment", editor.doc[2].get_text())
                self.assertEqual(editor.file_manifest[0]["output_end_page"], 2)
                self.assertEqual(editor.file_manifest[1]["output_start_page"], 3)
            finally:
                editor.doc.close()


if __name__ == "__main__":
    unittest.main()
