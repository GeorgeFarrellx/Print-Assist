from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from . import APP_NAME

PRINTABLE_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".xlsb", ".msg"}
ZIP_EXTENSIONS = {".zip"}
SUPPORTED_EXTENSIONS = PRINTABLE_EXTENSIONS | ZIP_EXTENSIONS


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


def get_supported_files_from_folder(folder_path: Path) -> tuple[list[Path], list[Path]]:
    supported: list[Path] = []
    unsupported: list[Path] = []
    for child in folder_path.iterdir():
        if not child.is_file():
            continue
        if child.suffix.lower() in SUPPORTED_EXTENSIONS:
            supported.append(child)
        else:
            unsupported.append(child)
    supported.sort(key=lambda p: p.name.lower())
    unsupported.sort(key=lambda p: p.name.lower())
    return supported, unsupported



def get_supported_files_from_client_folder(folder_path: Path) -> tuple[list[Path], list[Path]]:
    supported: list[Path] = []
    unsupported: list[Path] = []
    attachment_folder_names = {"attachments", "attachment", "email attachments"}

    direct_files: list[Path] = []
    attachment_files: list[Path] = []

    children = sorted(folder_path.iterdir(), key=lambda p: p.name.lower())
    for child in children:
        if child.is_file():
            if child.suffix.lower() in SUPPORTED_EXTENSIONS:
                direct_files.append(child)
            else:
                unsupported.append(child)
            continue

        if child.is_dir() and child.name.lower() in attachment_folder_names:
            for nested in sorted(child.iterdir(), key=lambda p: p.name.lower()):
                if not nested.is_file():
                    continue
                if nested.suffix.lower() in SUPPORTED_EXTENSIONS:
                    attachment_files.append(nested)
                else:
                    unsupported.append(nested)

    supported.extend(direct_files)
    supported.extend(attachment_files)
    unsupported.sort(key=lambda p: p.name.lower())
    return supported, unsupported

def default_output_path(files: list[Path]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{APP_NAME} - {timestamp}.pdf"
    if files:
        parents = {f.parent for f in files}
        if len(parents) == 1:
            return next(iter(parents)) / filename
    return Path.cwd() / filename
