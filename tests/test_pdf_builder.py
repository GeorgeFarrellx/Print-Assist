from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz
from PIL import Image

from print_assist.pdf_builder import (
    A4_PORTRAIT,
    _add_image_page,
    _add_pdf_pages,
    build_combined_pdf,
)


def _create_pdf(path: Path, width: float, height: float) -> None:
    doc = fitz.open()
    doc.new_page(width=width, height=height)
    doc.save(path)
    doc.close()


class _FakeOfficeConversionSession:
    def __init__(self, temp_dir: Path) -> None:
        self.temp_dir = temp_dir

    def __enter__(self) -> _FakeOfficeConversionSession:
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> None:
        _ = exc_type, exc_value, traceback

    def convert(self, source_path: Path) -> Path:
        _ = source_path
        output_path = self.temp_dir / "message.converted.pdf"
        _create_pdf(output_path, width=400, height=200)
        return output_path


class PdfBuilderTests(unittest.TestCase):
    def test_landscape_pdf_page_can_be_rotated_while_preserving_page_size(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "landscape.pdf"
            _create_pdf(source_path, width=400, height=200)

            output = fitz.open()
            try:
                _add_pdf_pages(
                    output,
                    source_path,
                    preserve_page_size=True,
                    force_portrait=True,
                )

                self.assertEqual(output[0].rect.width, 200)
                self.assertEqual(output[0].rect.height, 400)
            finally:
                output.close()

    def test_landscape_image_is_rotated_onto_portrait_a4_page(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "landscape.png"
            Image.new("RGB", (400, 200), "white").save(image_path)

            output = fitz.open()
            try:
                _add_image_page(output, image_path)

                self.assertAlmostEqual(output[0].rect.width, A4_PORTRAIT[0], places=3)
                self.assertAlmostEqual(output[0].rect.height, A4_PORTRAIT[1], places=3)
                self.assertLess(output[0].rect.width, output[0].rect.height)
            finally:
                output.close()

    def test_msg_converted_landscape_pages_are_added_as_portrait_pages(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            source_path = temp_path / "message.msg"
            output_path = temp_path / "combined.pdf"
            source_path.write_bytes(b"dummy msg")

            with patch(
                "print_assist.pdf_builder.OfficeConversionSession",
                _FakeOfficeConversionSession,
            ):
                processed, warnings = build_combined_pdf([source_path], output_path)

            self.assertEqual(processed, [str(source_path)])
            self.assertEqual(warnings, [])

            doc = fitz.open(output_path)
            try:
                self.assertEqual(doc[0].rect.width, 200)
                self.assertEqual(doc[0].rect.height, 400)
            finally:
                doc.close()


if __name__ == "__main__":
    unittest.main()
