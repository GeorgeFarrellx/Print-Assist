from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from print_assist.office_converter import (
    OfficeConversionSession,
    _memo_heading_from_message,
    convert_to_pdf_if_needed,
)
from print_assist.pdf_builder import _add_pdf_pages


class OfficeConverterTests(unittest.TestCase):
    def test_memo_heading_uses_first_message_recipient(self) -> None:
        message = type("Message", (), {"To": "Parkers Accountancy; George Farrell"})()

        self.assertEqual(_memo_heading_from_message(message), "Parkers Accountancy")

    def test_memo_heading_has_safe_fallback(self) -> None:
        message = type("Message", (), {"To": ""})()

        self.assertEqual(_memo_heading_from_message(message), "Outlook Email")

    def test_batch_conversion_uses_unique_output_names(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            session = OfficeConversionSession(Path(temp_dir))
            with patch.object(session, "_msg_to_pdf") as converter:
                first = session.convert(Path("first") / "message.msg")
                second = session.convert(Path("second") / "message.msg")

        self.assertEqual(first.name, "0001_message.converted.pdf")
        self.assertEqual(second.name, "0002_message.converted.pdf")
        self.assertNotEqual(first, second)
        self.assertEqual(converter.call_count, 2)

    def test_msg_conversion_uses_unattended_outlook_converter(self) -> None:
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
