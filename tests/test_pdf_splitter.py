from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from print_assist.pdf_splitter import (
    PdfPart,
    build_split_plan,
    export_split_pdf,
    inspect_pdf,
    parse_page_ranges,
    part_filename,
    safe_output_prefix,
)


def create_pdf(path: Path, page_count: int = 6) -> None:
    with fitz.open() as document:
        for index in range(page_count):
            page = document.new_page(width=400 + index * 10, height=600)
            page.insert_text((40, 40), f"Source page {index + 1}")
        document[1].set_rotation(90)
        document[2].set_cropbox(fitz.Rect(10, 10, 410, 590))
        document[0].add_text_annot((100, 100), "Keep this note")
        document[0].insert_link({"kind": fitz.LINK_URI, "from": fitz.Rect(40, 30, 140, 45), "uri": "https://example.com"})
        document.set_metadata({"title": "Synthetic statement fixture"})
        document.save(path)


class PdfSplitterTests(unittest.TestCase):
    def test_reading_page_count_does_not_inspect_text_or_choose_splits(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "statements.pdf"
            create_pdf(source)
            with patch.object(fitz.Page, "get_text", side_effect=AssertionError("No text analysis")):
                info = inspect_pdf(source)
            self.assertEqual(info.page_count, 6)
            self.assertEqual(build_split_plan(info.page_count, ()), (PdfPart(1, 6),))

    def test_user_marked_starts_cover_every_page_once(self) -> None:
        self.assertEqual(build_split_plan(10, (7, 4)), (PdfPart(1, 3), PdfPart(4, 6), PdfPart(7, 10)))

    def test_last_page_can_start_a_single_page_file(self) -> None:
        parts = build_split_plan(5, (5,))
        self.assertEqual(parts, (PdfPart(1, 4), PdfPart(5, 5)))
        self.assertEqual(parts[-1].page_range, "5")

    def test_marked_starts_reject_out_of_bounds_pages(self) -> None:
        for starts in ((0,), (7,), (-1,), (True,), (1.5,)):
            with self.subTest(starts=starts), self.assertRaises(ValueError):
                build_split_plan(6, starts)

    def test_typed_ranges_support_spaces_single_pages_and_semicolons(self) -> None:
        self.assertEqual(parse_page_ranges(" 1 - 3; 4; 5-6 ", 6), (PdfPart(1, 3), PdfPart(4, 4), PdfPart(5, 6)))
        self.assertEqual(parse_page_ranges("1\u20133, 4\u20136", 6), (PdfPart(1, 3), PdfPart(4, 6)))

    def test_typed_ranges_reject_missing_duplicate_reversed_or_invalid_pages(self) -> None:
        for ranges in ("", "1-3", "2-6", "1-3, 5-6", "1-4, 4-6", "4-6, 1-3", "1-7", "1-0, 1-6", "1-3, 6-4", "1-3,", "one-six", "1.5-6"):
            with self.subTest(ranges=ranges), self.assertRaises(ValueError):
                parse_page_ranges(ranges, 6)

    def test_export_keeps_text_page_geometry_rotation_annotations_and_original(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "source.pdf"
            create_pdf(source)
            original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
            parts = parse_page_ranges("1-3, 4-6", 6)
            progress = []
            result = export_split_pdf(source, parts, folder, "Statements", lambda current, total: progress.append((current, total)))
            self.assertEqual(progress, [(1, 2), (2, 2)])
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), original_hash)
            self.assertEqual(result.folder.name, "Statements - split")
            with fitz.open(source) as original:
                source_index = 0
                for file_path, part in zip(result.files, parts):
                    with fitz.open(file_path) as output:
                        self.assertEqual(len(output), part.page_count)
                        self.assertEqual(output.metadata["title"], original.metadata["title"])
                        for page in output:
                            source_page = original[source_index]
                            self.assertEqual(page.get_text(), source_page.get_text())
                            self.assertEqual(page.rect, source_page.rect)
                            self.assertEqual(page.mediabox, source_page.mediabox)
                            self.assertEqual(page.cropbox, source_page.cropbox)
                            self.assertEqual(page.rotation, source_page.rotation)
                            self.assertEqual(page.get_pixmap().samples, source_page.get_pixmap().samples)
                            source_index += 1
                self.assertEqual(source_index, len(original))
            with fitz.open(result.files[0]) as output:
                self.assertEqual(output[0].first_annot.info["content"], "Keep this note")
                self.assertEqual(output[0].get_links()[0]["uri"], "https://example.com")

    def test_blank_and_scanned_pages_are_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "scan.pdf"
            with fitz.open() as document:
                document.new_page()
                page = document.new_page()
                pixmap = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 30, 20), False)
                pixmap.clear_with(125)
                page.insert_image(fitz.Rect(30, 30, 330, 230), pixmap=pixmap)
                document.save(source)
            result = export_split_pdf(source, (PdfPart(1, 1), PdfPart(2, 2)), folder, "Scan")
            with fitz.open(result.files[0]) as blank, fitz.open(result.files[1]) as scan:
                self.assertEqual(blank.page_count, 1)
                self.assertEqual(scan.page_count, 1)
                self.assertEqual(scan[0].get_text(), "")
                self.assertEqual(len(scan[0].get_images()), 1)

    def test_repeat_export_uses_new_folder_without_overwriting_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "source.pdf"
            create_pdf(source)
            parts = build_split_plan(6, (4,))
            first = export_split_pdf(source, parts, folder, "Statements")
            hashes = [hashlib.sha256(path.read_bytes()).hexdigest() for path in first.files]
            second = export_split_pdf(source, parts, folder, "Statements")
            self.assertEqual(second.folder.name, "Statements - split (2)")
            self.assertEqual([hashlib.sha256(path.read_bytes()).hexdigest() for path in first.files], hashes)

    def test_failed_export_removes_partial_files_and_preserves_existing_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "source.pdf"
            create_pdf(source)
            existing = folder / "Statements - split"
            existing.mkdir()
            (existing / "keep.txt").write_text("Existing file", encoding="utf-8")

            def fail_after_first_file(current: int, total: int) -> None:
                raise OSError("Simulated disk failure")

            with self.assertRaisesRegex(OSError, "Simulated disk failure"):
                export_split_pdf(source, build_split_plan(6, (4,)), folder, "Statements", fail_after_first_file)
            self.assertEqual(sorted(path.name for path in folder.iterdir()), ["Statements - split", "source.pdf"])
            self.assertEqual((existing / "keep.txt").read_text(encoding="utf-8"), "Existing file")

    def test_incomplete_export_plan_is_rejected_before_creating_a_folder(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "source.pdf"
            create_pdf(source)
            for parts in ((), (PdfPart(1, 3),), (PdfPart(1, 4), PdfPart(4, 6))):
                with self.subTest(parts=parts), self.assertRaises(ValueError):
                    export_split_pdf(source, parts, folder, "Statements")
            self.assertEqual(list(folder.iterdir()), [source])

    def test_changed_source_is_rejected_before_export(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "source.pdf"
            create_pdf(source)
            info = inspect_pdf(source)
            with self.assertRaisesRegex(ValueError, "source PDF changed"):
                export_split_pdf(source, build_split_plan(6, (4,)), folder, "Statements", expected_signature=(0, info.source_signature[1]))
            self.assertEqual(list(folder.iterdir()), [source])

    def test_password_protected_pdf_gives_actionable_error(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "locked.pdf"
            with fitz.open() as document:
                document.new_page()
                document.save(source, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw="owner", user_pw="secret")
            with self.assertRaisesRegex(ValueError, "unlocked copy"):
                inspect_pdf(source)

    def test_output_names_cannot_escape_the_new_folder(self) -> None:
        prefix = '../../Statements:2025? '
        filename = part_filename(prefix, 1, PdfPart(1, 5))
        self.assertEqual(Path(filename).name, filename)
        self.assertNotIn(":", filename)
        self.assertNotIn("?", filename)
        self.assertEqual(safe_output_prefix(". "), "PDF")
        self.assertLessEqual(len(safe_output_prefix("a" * 300)), 70)


if __name__ == "__main__":
    unittest.main()
