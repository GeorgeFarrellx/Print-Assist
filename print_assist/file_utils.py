from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from . import APP_NAME

SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".xlsb", ".msg"}


def filter_supported_files(paths: Iterable[str]) -> tuple[list[Path], list[Path]]:
    supported: list[Path] = []
    unsupported: list[Path] = []
    for raw_path in paths:
        path = Path(raw_path)
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            supported.append(path)
        else:
            unsupported.append(path)
    return supported, unsupported


def default_output_path(files: list[Path]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{APP_NAME} - {timestamp}.pdf"
    if files:
        parents = {f.parent for f in files}
        if len(parents) == 1:
            return next(iter(parents)) / filename
    return Path.cwd() / filename
