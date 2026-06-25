from __future__ import annotations

from datetime import datetime
from pathlib import Path


_INVALID_WINDOWS_FILENAME_CHARS = set('<>:"/\\|?*')
_RESERVED_WINDOWS_FILENAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


def safe_outlook_attachment_name(raw_name: str, fallback: str) -> str:
    basename = str(raw_name or "").replace("\\", "/").rsplit("/", 1)[-1]
    cleaned = "".join(
        "_" if char in _INVALID_WINDOWS_FILENAME_CHARS or ord(char) < 32 else char
        for char in basename
    )
    cleaned = cleaned.strip().rstrip(" .")
    if cleaned in {"", ".", ".."}:
        cleaned = fallback

    stem = Path(cleaned).stem or cleaned
    if stem.upper() in _RESERVED_WINDOWS_FILENAMES:
        cleaned = f"_{cleaned}"

    if len(cleaned) > 180:
        suffix = Path(cleaned).suffix
        stem = Path(cleaned).stem or fallback
        cleaned = f"{stem[: max(1, 180 - len(suffix))]}{suffix}"
    return cleaned


def unique_file_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _attachment_is_hidden(attachment: object) -> bool:
    try:
        property_accessor = attachment.PropertyAccessor
        return bool(property_accessor.GetProperty(_PR_ATTACHMENT_HIDDEN))
    except Exception:
        return False


def _add_detected_extension(path: Path) -> Path:
    if path.suffix.lower() not in {"", ".bin"}:
        return path

    try:
        with path.open("rb") as source:
            header = source.read(16)
    except OSError:
        return path

    detected_suffix = ""
    if header.startswith(b"%PDF"):
        detected_suffix = ".pdf"
    elif header.startswith(b"\xff\xd8\xff"):
        detected_suffix = ".jpg"
    elif header.startswith(b"\x89PNG\r\n\x1a\n"):
        detected_suffix = ".png"
    elif header.startswith(b"BM"):
        detected_suffix = ".bmp"
    elif header.startswith((b"II*\x00", b"MM\x00*")):
        detected_suffix = ".tiff"

    if not detected_suffix:
        return path

    detected_path = unique_file_path(path.with_suffix(detected_suffix))
    path.replace(detected_path)
    return detected_path


def _save_visible_attachments(message: object, target_dir: Path) -> tuple[list[Path], list[str]]:
    target_dir.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    warnings: list[str] = []
    attachments = message.Attachments

    for index in range(1, attachments.Count + 1):
        attachment = None
        try:
            attachment = attachments.Item(index)
            if _attachment_is_hidden(attachment):
                continue

            raw_name = getattr(attachment, "FileName", "") or getattr(attachment, "DisplayName", "")
            safe_name = safe_outlook_attachment_name(raw_name, f"attachment_{index}.bin")
            output_path = unique_file_path(target_dir / safe_name)
            attachment.SaveAsFile(str(output_path))
            saved.append(_add_detected_extension(output_path))
        except Exception as exc:
            display_name = getattr(attachment, "FileName", "") if attachment is not None else ""
            label = display_name or f"attachment {index}"
            warnings.append(f"Could not extract '{label}': {exc}")
        finally:
            attachment = None

    return saved, warnings


def _message_datetime(message: object) -> datetime | None:
    for attribute_name in ("SentOn", "ReceivedTime", "CreationTime"):
        try:
            value = getattr(message, attribute_name)
        except Exception:
            continue
        if isinstance(value, datetime):
            return value
    return None


def extract_msg_details(
    source_path: Path,
    target_dir: Path,
) -> tuple[list[Path], list[str], datetime | None]:
    try:
        import pythoncom
        import win32com.client
    except Exception as exc:
        raise RuntimeError("Outlook/pywin32 automation is unavailable on this system") from exc

    pythoncom.CoInitialize()
    outlook = None
    namespace = None
    message = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        message = namespace.OpenSharedItem(str(source_path))
        attachments, warnings = _save_visible_attachments(message, target_dir)
        return attachments, warnings, _message_datetime(message)
    except Exception as exc:
        raise RuntimeError(f"Failed to read attachments from Outlook message: {source_path.name}") from exc
    finally:
        message = None
        namespace = None
        outlook = None
        pythoncom.CoUninitialize()


def extract_msg_attachments(source_path: Path, target_dir: Path) -> tuple[list[Path], list[str]]:
    attachments, warnings, _ = extract_msg_details(source_path, target_dir)
    return attachments, warnings
