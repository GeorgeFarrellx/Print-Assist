from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from print_assist.office_converter import convert_to_pdf_if_needed
from print_assist.pdf_builder import _add_pdf_pages


class OfficeConverterTests(unittest.TestCase):
    def test_msg_conversion_uses_outlook_memo_style_printer(self) -> None:
        source = Path("message.msg")

        with tempfile.TemporaryDirectory() as temp_dir, patch(
            "print_assist.office_converter._msg_to_pdf"
        ) as converter:
            output = convert_to_pdf_if_needed(source, Path(temp_dir))

        converter.assert_called_once_with(source, output, Path(temp_dir))
        self.assertEqual(output.name, "message.converted.pdf")

    def test_outlook_pdf_page_size_can_be_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "memo.pdf"
            source = fitz.open()
            source.new_page(width=612, height=792)
            source.save(source_path)
            source.close()

            output = fitz.open()
            _add_pdf_pages(output, source_path, preserve_page_size=True)

            self.assertEqual(output[0].rect.width, 612)
            self.assertEqual(output[0].rect.height, 792)
            output.close()


if __name__ == "__main__":
    unittest.main()
