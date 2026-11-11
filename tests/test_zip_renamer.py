from __future__ import annotations

import csv
import hashlib
import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch

from print_assist.app import PrintAssistApp
from print_assist.zip_renamer import (
    ZipExtractionPathError,
    ZipExtractionWarning,
    _windows_path_length,
    default_extracted_folder_path,
    rename_and_extract_zip_contents,
)


def make_zip(path: Path, members: list[tuple[str, bytes]]) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in members:
            archive.writestr(name, data)
    return path


class ZipRenamerTests(unittest.TestCase):
    def test_extraction_keeps_document_bytes_folders_and_name_mapping(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = make_zip(root / "batch.zip", [
                ("client/Attachments/invoice.docx", b"original document bytes"),
                ("client/statement.pdf", b"original PDF bytes"),
            ])
            original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
            output = rename_and_extract_zip_contents(source, root / "extracted")

            self.assertEqual((output / "client/Attachments/A1.docx").read_bytes(), b"original document bytes")
            self.assertEqual((output / "client/A2.pdf").read_bytes(), b"original PDF bytes")
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), original_hash)
            with (output / "Name_Mapping.csv").open(newline="", encoding="utf-8") as mapping:
                rows = [row for row in csv.DictReader(mapping) if row["RowType"] == "File"]
            self.assertEqual(rows[0]["OriginalPath"], "client/Attachments/invoice.docx")
            self.assertEqual(rows[0]["NewPath"], "client/Attachments/A1.docx")

    def test_overlong_later_file_is_rejected_before_any_extraction(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            output = root / "extracted"
            nested = "x" * (260 - len(str(output.resolve())) - len("/A2.docx") - 1)
            source = make_zip(root / "batch.zip", [
                ("first.pdf", b"first"),
                (f"{nested}/invoice.docx", b"second"),
            ])

            with self.assertRaises(ZipExtractionPathError) as caught:
                rename_and_extract_zip_contents(source, output)

            self.assertFalse(output.exists())
            self.assertEqual(caught.exception.longest_path.name, "A2.docx")
            self.assertIn("260 characters; maximum 259", str(caught.exception))

    def test_file_at_exact_windows_limit_can_be_extracted(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            output = root / "extracted"
            nested = "x" * (259 - len(str(output.resolve())) - len("/A1.docx") - 1)
            source = make_zip(root / "batch.zip", [(f"{nested}/invoice.docx", b"document")])

            rename_and_extract_zip_contents(source, output)

            document = output / nested / "A1.docx"
            self.assertEqual(len(str(document.resolve())), 259)
            self.assertEqual(document.read_bytes(), b"document")

    def test_collision_suffix_is_included_in_preflight_length(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            output = root / "extracted"
            nested = "x" * (259 - len(str(output.resolve())) - len("/A1.docx") - 1)
            source = make_zip(root / "batch.zip", [
                (f"{nested}/A1.docx/", b""),
                (f"{nested}/invoice.docx", b"document"),
            ])

            with self.assertRaises(ZipExtractionPathError) as caught:
                rename_and_extract_zip_contents(source, output)

            self.assertEqual(caught.exception.longest_path.name, "A1 (2).docx")
            self.assertFalse(output.exists())

    def test_mapping_path_is_checked_even_for_an_empty_zip(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            output = root / ("x" * (260 - len(str(root.resolve())) - len("/Name_Mapping.csv") - 1))
            source = make_zip(root / "batch.zip", [])

            with self.assertRaises(ZipExtractionPathError) as caught:
                rename_and_extract_zip_contents(source, output)

            self.assertEqual(caught.exception.longest_path.name, "Name_Mapping.csv")
            self.assertFalse(output.exists())

    def test_windows_limit_counts_utf16_units_for_unicode_folders(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            output = root / "extracted"
            nested = "x" * (259 - len(str(output.resolve())) - len("/A1.docx") - 2) + "\U0001f4c1"
            source = make_zip(root / "batch.zip", [(f"{nested}/invoice.docx", b"document")])
            self.assertEqual(len(str(output / nested / "A1.docx")), 259)

            with self.assertRaises(ZipExtractionPathError):
                rename_and_extract_zip_contents(source, output)
            self.assertFalse(output.exists())

    @unittest.skipUnless(os.name == "nt", "Windows path namespaces")
    def test_extended_local_and_unc_prefixes_do_not_hide_real_path_length(self) -> None:
        for regular, extended in (
            (r"C:\folder\file.docx", r"\\?\C:\folder\file.docx"),
            (r"\\server\share\file.docx", r"\\?\UNC\server\share\file.docx"),
        ):
            with self.subTest(path=extended):
                self.assertEqual(_windows_path_length(Path(extended)), len(regular))

    def test_unsafe_entries_still_warn_and_do_not_change_safe_file_sequence(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = make_zip(root / "batch.zip", [
                ("../outside.pdf", b"unsafe"),
                ("client/invoice.docx", b"safe"),
            ])
            output = root / "extracted"

            with self.assertRaises(ZipExtractionWarning):
                rename_and_extract_zip_contents(source, output)

            self.assertEqual((output / "client/A1.docx").read_bytes(), b"safe")
            self.assertFalse((root / "outside.pdf").exists())


class ZipExtractionDialogTests(unittest.TestCase):
    def app_for(self, sources: list[Path]) -> PrintAssistApp:
        app = PrintAssistApp.__new__(PrintAssistApp)
        app.files = sources
        app._selected_tree_paths = Mock(return_value=[])
        app.status_var = Mock()
        return app

    def test_long_destination_can_be_retried_with_a_shorter_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            long_parent = root / ("x" * (230 - len(str(root.resolve())) - 1))
            long_parent.mkdir()
            source = make_zip(root / "batch.zip", [("client/Attachments/invoice.docx", b"document")])
            app = self.app_for([source])

            with patch("print_assist.app.filedialog.askdirectory", side_effect=[str(long_parent), str(root)]) as picker, patch("print_assist.app.messagebox") as messages:
                app.rename_zip_contents()

            self.assertEqual(picker.call_count, 2)
            messages.showerror.assert_called_once()
            self.assertIn("Choose an extraction folder closer", messages.showerror.call_args.args[1])
            self.assertFalse(default_extracted_folder_path(source, long_parent).exists())
            output = default_extracted_folder_path(source, root)
            self.assertEqual((output / "client/Attachments/A1.docx").read_bytes(), b"document")
            messages.showinfo.assert_called_once()

    def test_cancelling_retry_keeps_completed_zip_outputs_and_stops_batch(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            parent = root / ("x" * (180 - len(str(root.resolve())) - 1))
            parent.mkdir()
            first = make_zip(root / "first.zip", [("invoice.docx", b"first")])
            second = make_zip(root / "second.zip", [(f"{'nested' * 15}/invoice.docx", b"second")])
            third = make_zip(root / "third.zip", [("invoice.docx", b"third")])
            app = self.app_for([first, second, third])

            with patch("print_assist.app.filedialog.askdirectory", side_effect=[str(parent), ""]), patch("print_assist.app.messagebox"):
                app.rename_zip_contents()

            self.assertEqual((default_extracted_folder_path(first, parent) / "A1.docx").read_bytes(), b"first")
            self.assertFalse(default_extracted_folder_path(second, parent).exists())
            self.assertFalse(default_extracted_folder_path(third, parent).exists())
            self.assertIn("cancelled", app.status_var.set.call_args.args[0])


if __name__ == "__main__":
    unittest.main()
