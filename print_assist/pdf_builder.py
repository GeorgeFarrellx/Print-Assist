from __future__ import annotations

from pathlib import Path

import fitz
from PIL import Image

A4_PORTRAIT = (595.2756, 841.8898)
A4_LANDSCAPE = (841.8898, 595.2756)
MARGIN = 24


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


def _add_pdf_pages(output_doc: fitz.Document, source_path: Path) -> None:
    with fitz.open(source_path) as src_doc:
        for src_page in src_doc:
            src_rect = src_page.rect
            page_w, page_h = _choose_a4_orientation(src_rect.width, src_rect.height)
            out_page = output_doc.new_page(width=page_w, height=page_h)
            target = _fit_rect(src_rect.width, src_rect.height, page_w, page_h, MARGIN)
            out_page.show_pdf_page(target, src_doc, src_page.number)


def _add_image_page(output_doc: fitz.Document, image_path: Path) -> None:
    with Image.open(image_path) as img:
        width, height = img.size

    page_w, page_h = _choose_a4_orientation(width, height)
    out_page = output_doc.new_page(width=page_w, height=page_h)
    target = _fit_rect(width, height, page_w, page_h, MARGIN)
    out_page.insert_image(target, filename=str(image_path), keep_proportion=True)


def build_combined_pdf(files: list[Path], output_path: Path) -> tuple[list[str], list[str]]:
    processed: list[str] = []
    warnings: list[str] = []
    output_doc = fitz.open()
    try:
        for file_path in files:
            try:
                suffix = file_path.suffix.lower()
                if suffix == ".pdf":
                    _add_pdf_pages(output_doc, file_path)
                    processed.append(str(file_path))
                else:
                    _add_image_page(output_doc, file_path)
                    processed.append(str(file_path))
            except Exception as exc:
                warnings.append(f"Could not process '{file_path.name}': {exc}")

        if not processed:
            raise ValueError("No files were successfully processed.")

        output_doc.save(output_path)
        return processed, warnings
    finally:
        output_doc.close()
