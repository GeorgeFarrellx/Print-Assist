from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from print_assist.app import PrintAssistApp, format_file_selection_summary, reorder_grouped_files
from print_assist.outlook_message import _save_visible_attachments


class _PropertyAccessor:
    def __init__(self, hidden: bool) -> None:
        self.hidden = hidden

    def GetProperty(self, property_name: str) -> bool:
        _ = property_name
        return self.hidden


class _Attachment:
    def __init__(self, filename: str, payload: bytes, hidden: bool = False) -> None:
        self.FileName = filename
        self.DisplayName = filename
        self.payload = payload
        self.PropertyAccessor = _PropertyAccessor(hidden)

    def SaveAsFile(self, output_path: str) -> None:
        Path(output_path).write_bytes(self.payload)


class _Attachments:
    def __init__(self, attachments: list[_Attachment]) -> None:
        self._attachments = attachments
        self.Count = len(attachments)

    def Item(self, index: int) -> _Attachment:
        return self._attachments[index - 1]


class OutlookMessageTests(unittest.TestCase):
    def test_file_summary_separates_emails_and_extracted_attachments(self) -> None:
        outlook_temp_dir = Path("temp") / "outlook"
        files = [
            *(Path(f"email_{index}.msg") for index in range(5)),
            *(outlook_temp_dir / "message_attachments" / f"attachment_{index}.pdf" for index in range(4)),
        ]

        summary = format_file_selection_summary(files, outlook_temp_dir)

        self.assertEqual(summary, "Selected: 5 emails + 4 attachments")

    def test_file_summary_includes_other_files_for_mixed_selection(self) -> None:
        outlook_temp_dir = Path("temp") / "outlook"
        files = [
            Path("email.msg"),
            outlook_temp_dir / "message_attachments" / "attachment.pdf",
            Path("manually-added.pdf"),
        ]

        summary = format_file_selection_summary(files, outlook_temp_dir)

        self.assertEqual(summary, "Selected: 1 email + 1 attachment + 1 other file")

    def test_file_summary_stays_simple_without_outlook_items(self) -> None:
        summary = format_file_selection_summary([Path("one.pdf"), Path("two.jpg")])

        self.assertEqual(summary, "Selected files: 2")

    def test_visible_attachments_are_saved_and_hidden_inline_images_are_skipped(self) -> None:
        message = SimpleNamespace(
            Attachments=_Attachments(
                [
                    _Attachment("statement.pdf", b"%PDF-1.7"),
                    _Attachment("signature.png", b"\x89PNG\r\n\x1a\n", hidden=True),
                ]
            )
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            saved, warnings = _save_visible_attachments(message, Path(temp_dir))

            self.assertEqual(warnings, [])
            self.assertEqual([path.name for path in saved], ["statement.pdf"])
            self.assertEqual(saved[0].read_bytes(), b"%PDF-1.7")

    def test_attachment_without_extension_gets_detected_pdf_extension(self) -> None:
        message = SimpleNamespace(
            Attachments=_Attachments([_Attachment("statement", b"%PDF-1.7\ncontent")])
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            saved, warnings = _save_visible_attachments(message, Path(temp_dir))

            self.assertEqual(warnings, [])
            self.assertEqual(saved[0].name, "statement.pdf")

    def test_msg_path_expands_email_then_printable_attachments_in_order(self) -> None:
        source_message = Path("email.msg")
        nested_message = Path("attached email.msg")
        first_attachment = Path("first.pdf")
        nested_attachment = Path("nested.jpg")
        unsupported_attachment = Path("notes.txt")
        fake_app = SimpleNamespace(
            _get_outlook_drop_temp_dir=lambda: Path("temp"),
        )

        def fake_extract(path: Path, target_dir: Path) -> tuple[list[Path], list[str]]:
            _ = target_dir
            if path == source_message:
                return [first_attachment, nested_message, unsupported_attachment], []
            if path == nested_message:
                return [nested_attachment], []
            raise AssertionError(f"Unexpected message: {path}")

        with patch("print_assist.app.extract_msg_attachments", side_effect=fake_extract):
            expanded, unsupported, warnings = PrintAssistApp._expand_outlook_message_paths(
                fake_app,
                [source_message],
            )

        self.assertEqual(
            expanded,
            [source_message, first_attachment, nested_message, nested_attachment],
        )
        self.assertEqual(unsupported, [unsupported_attachment])
        self.assertEqual(warnings, [])

    def test_msg_expansion_keeps_each_attachment_with_its_email(self) -> None:
        source_message = Path("email.msg")
        nested_message = Path("attached email.msg")
        first_attachment = Path("first.pdf")
        nested_attachment = Path("nested.jpg")
        fake_app = SimpleNamespace(
            _get_outlook_drop_temp_dir=lambda: Path("temp"),
        )

        def fake_extract(path: Path, target_dir: Path) -> tuple[list[Path], list[str]]:
            _ = target_dir
            if path == source_message:
                return [first_attachment, nested_message], []
            if path == nested_message:
                return [nested_attachment], []
            raise AssertionError(f"Unexpected message: {path}")

        with patch("print_assist.app.extract_msg_attachments", side_effect=fake_extract):
            expanded, parents, unsupported, warnings = (
                PrintAssistApp._expand_outlook_message_entries(fake_app, [source_message])
            )

        self.assertEqual(
            expanded,
            [source_message, first_attachment, nested_message, nested_attachment],
        )
        self.assertEqual(
            parents,
            {
                source_message: None,
                first_attachment: source_message,
                nested_message: source_message,
                nested_attachment: nested_message,
            },
        )
        self.assertEqual(unsupported, [])
        self.assertEqual(warnings, [])

    def test_reordering_email_moves_its_attachment_group_as_one_unit(self) -> None:
        email = Path("email.msg")
        attachment = Path("attachment.pdf")
        other = Path("other.pdf")
        files = [email, attachment, other]
        parents = {email: None, attachment: email, other: None}

        reordered = reorder_grouped_files(files, parents, {email}, 1)

        self.assertEqual(reordered, [other, email, attachment])


if __name__ == "__main__":
    unittest.main()
