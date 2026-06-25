from __future__ import annotations

import tempfile
from pathlib import Path

import fitz
from PIL import Image

from .office_converter import (
    EXCEL_EXTENSIONS,
    MSG_EXTENSIONS,
    WORD_EXTENSIONS,
    OfficeConversionSession,
)

A4_PORTRAIT = (595.2756, 841.8898)
A4_LANDSCAPE = (841.8898, 595.2756)
MARGIN = 24
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
OFFICE_EXTENSIONS = WORD_EXTENSIONS | EXCEL_EXTENSIONS | MSG_EXTENSIONS


def _fit_rect(src_width: float, src_height: float, page_width: float, page_height: float, margin: float) -> fitz.Rect:
    usable_w = page_width - (2 * margin)
    usable_h = page_height - (2 * margin)
    scale = min(usable_w / src_width, usable_h / src_height)
    draw_w = src_width * scale
    draw_h = src_height * scale
    x0 = (page_width - draw_w) / 2
    y0 = (page_height - draw_h) / 2
    return fitz.Rect(x0, y0, x0 + draw_w, y0 + draw_h)


def _choose_a4_orientation(width: float, height: float) -> tuple[float, float]:
    return A4_LANDSCAPE if width > height else A4_PORTRAIT


def _add_pdf_pages(
    output_doc: fitz.Document,
    source_path: Path,
    preserve_page_size: bool = False,
) -> None:
    with fitz.open(source_path) as src_doc:
        for src_page in src_doc:
            src_rect = src_page.rect
            if preserve_page_size:
                page_w, page_h = src_rect.width, src_rect.height
            else:
                page_w, page_h = _choose_a4_orientation(src_rect.width, src_rect.height)
            out_page = output_doc.new_page(width=page_w, height=page_h)
            target = (
                out_page.rect
                if preserve_page_size
                else _fit_rect(src_rect.width, src_rect.height, page_w, page_h, MARGIN)
            )
            out_page.show_pdf_page(target, src_doc, src_page.number)


def _add_image_page(output_doc: fitz.Document, image_path: Path) -> None:
    with Image.open(image_path) as img:
        width, height = img.size

    page_w, page_h = _choose_a4_orientation(width, height)
    out_page = output_doc.new_page(width=page_w, height=page_h)
    target = _fit_rect(width, height, page_w, page_h, MARGIN)
    out_page.insert_image(target, filename=str(image_path), keep_proportion=True)


def build_combined_pdf(
    files: list[Path],
    output_path: Path,
    progress_callback: callable | None = None,
    manifest_callback: callable | None = None,
) -> tuple[list[str], list[str]]:
    processed: list[str] = []
    warnings: list[str] = []
    with tempfile.TemporaryDirectory(prefix="print_assist_") as temp_dir_raw:
        temp_dir = Path(temp_dir_raw)
        output_doc = fitz.open()
        try:
            total_files = len(files)
            with OfficeConversionSession(temp_dir) as office_converter:
                for index, file_path in enumerate(files, start=1):
                    try:
                        suffix = file_path.suffix.lower()
                        start_page = len(output_doc) + 1
                        if suffix == ".pdf":
                            _add_pdf_pages(output_doc, file_path)
                        elif suffix in IMAGE_EXTENSIONS:
                            _add_image_page(output_doc, file_path)
                        elif suffix in OFFICE_EXTENSIONS:
                            converted_pdf = office_converter.convert(file_path)
                            _add_pdf_pages(
                                output_doc,
                                converted_pdf,
                                preserve_page_size=suffix in MSG_EXTENSIONS,
                            )
                        else:
                            raise ValueError(f"Unsupported file extension: {file_path.suffix}")
                        end_page = len(output_doc)
                        output_page_count = end_page - start_page + 1
                        processed.append(str(file_path))
                        if manifest_callback is not None:
                            manifest_callback(
                                {
                                    "source_path": str(file_path),
                                    "source_name": file_path.name,
                                    "source_extension": suffix,
                                    "output_start_page": start_page,
                                    "output_end_page": end_page,
                                    "output_page_count": output_page_count,
                                }
                            )
                        if progress_callback is not None:
                            progress_callback(index, total_files, file_path, f"Processing {index} of {total_files}: {file_path.name}")
                    except Exception as exc:
                        warnings.append(f"Could not process '{file_path.name}': {exc}")

            if not processed:
                raise ValueError("No files were successfully processed.")

            output_doc.save(output_path)
        finally:
            # Close the combined document before TemporaryDirectory cleanup.
            # PyMuPDF can retain source PDF handles until this point on Windows.
            output_doc.close()
    return processed, warnings
