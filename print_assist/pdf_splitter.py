from __future__ import annotations

import re
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path

import fitz


@dataclass(frozen=True)
class PdfInfo:
    page_count: int
    source_signature: tuple[int, int]


@dataclass(frozen=True)
class PdfPart:
    # All page numbers in the split plan are one-based and inclusive.
    start_page: int
    end_page: int

    @property
    def page_count(self) -> int:
        return self.end_page - self.start_page + 1

    @property
    def page_range(self) -> str:
        if self.start_page == self.end_page:
            return str(self.start_page)
        return f"{self.start_page}-{self.end_page}"


@dataclass(frozen=True)
class PdfSplitResult:
    folder: Path
    files: tuple[Path, ...]


def source_signature(source_path: Path) -> tuple[int, int]:
    stat = source_path.stat()
    return stat.st_size, stat.st_mtime_ns


def _validate_pdf(document: fitz.Document) -> None:
    if not document.is_pdf:
        raise ValueError("Choose a PDF file to split.")
    if document.needs_pass:
        raise ValueError("This PDF needs a password. Save an unlocked copy, then split that copy.")
    if document.page_count < 1:
        raise ValueError("This PDF has no pages to split.")


def inspect_pdf(source_path: Path) -> PdfInfo:
    """Read the page count only. Split positions are always supplied by the user."""
    signature = source_signature(source_path)
    with fitz.open(source_path) as document:
        _validate_pdf(document)
        page_count = document.page_count
    if source_signature(source_path) != signature:
        raise ValueError("The source PDF changed while it was being read. Reopen Split PDF.")
    return PdfInfo(page_count, signature)


def parse_page_ranges(value: str, page_count: int) -> tuple[PdfPart, ...]:
    """Read the user's ranges; require a complete, ordered partition of the PDF."""
    parts = []
    for token in re.split(r"[,;\n]", value.strip()):
        match = re.fullmatch(r"\s*([0-9]+)(?:\s*[-\u2013]\s*([0-9]+))?\s*", token)
        if match is None:
            raise ValueError("Enter page ranges separated by commas, for example: 1-5, 6-10, 11-20.")
        start = int(match[1])
        end = int(match[2]) if match[2] else start
        parts.append(PdfPart(start, end))
    validate_split_plan(tuple(parts), page_count)
    return tuple(parts)


def validate_split_plan(parts: tuple[PdfPart, ...], page_count: int) -> None:
    if type(page_count) is not int or page_count < 1:
        raise ValueError("The PDF must contain at least one page.")
    next_page = 1
    for part in parts:
        if (
            type(part.start_page) is not int
            or type(part.end_page) is not int
            or not 1 <= part.start_page <= part.end_page <= page_count
        ):
            raise ValueError(f"Each range must use pages between 1 and {page_count}, with the first page before the last.")
        if part.start_page != next_page:
            raise ValueError(f"The next range must start at page {next_page}. Include every page once, in order.")
        next_page = part.end_page + 1
    if not parts or next_page != page_count + 1:
        raise ValueError(f"Include every page once, ending at page {page_count}.")


def build_split_plan(
    page_count: int,
    starts: Iterable[int],
) -> tuple[PdfPart, ...]:
    if type(page_count) is not int or page_count < 1:
        raise ValueError("The PDF must contain at least one page.")
    raw_starts = tuple(starts)
    if any(type(page) is not int or not 1 <= page <= page_count for page in raw_starts):
        raise ValueError(f"Split pages must be between 1 and {page_count}.")
    ordered_starts = sorted({1, *raw_starts})
    parts = []
    for start, following in zip(ordered_starts, [*ordered_starts[1:], page_count + 1]):
        parts.append(PdfPart(start, following - 1))
    return tuple(parts)


def safe_output_prefix(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip().rstrip(". ")
    return cleaned[:70].rstrip(". ") or "PDF"


def part_filename(prefix: str, index: int, part: PdfPart) -> str:
    return f"{safe_output_prefix(prefix)} - {index:02d} - pages {part.page_range}.pdf"


def _create_output_folder(destination: Path, prefix: str) -> Path:
    base = f"{safe_output_prefix(prefix)} - split"
    for index in range(1, 10000):
        folder = destination / (base if index == 1 else f"{base} ({index})")
        try:
            folder.mkdir()
        except FileExistsError:
            continue
        return folder
    raise ValueError("Too many split folders already exist. Choose another output folder.")


def export_split_pdf(
    source_path: Path,
    parts: tuple[PdfPart, ...],
    destination: Path,
    prefix: str,
    progress_callback: Callable[[int, int], None] | None = None,
    expected_signature: tuple[int, int] | None = None,
) -> PdfSplitResult:
    """Copy every page exactly once into a new folder; roll back failed exports."""
    signature = source_signature(source_path)
    if expected_signature is not None and signature != expected_signature:
        raise ValueError("The source PDF changed. Reopen Split PDF before exporting.")
    folder: Path | None = None
    created: list[Path] = []
    try:
        with fitz.open(source_path) as source:
            _validate_pdf(source)
            validate_split_plan(parts, source.page_count)
            folder = _create_output_folder(destination, prefix)
            for index, part in enumerate(parts, start=1):
                output_path = folder / part_filename(prefix, index, part)
                created.append(output_path)
                with fitz.open() as output:
                    output.insert_pdf(source, from_page=part.start_page - 1, to_page=part.end_page - 1)
                    output.set_metadata(source.metadata)
                    output.save(output_path, garbage=4, deflate=True)
                if progress_callback is not None:
                    progress_callback(index, len(parts))
        if source_signature(source_path) != signature:
            raise ValueError("The source PDF changed during export. Reopen Split PDF.")
    except Exception:
        # Only remove files created by this export in its exclusively owned folder.
        if folder is not None:
            for output_path in created:
                output_path.unlink(missing_ok=True)
            folder.rmdir()
        raise
    return PdfSplitResult(folder, tuple(created))
