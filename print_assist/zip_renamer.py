from __future__ import annotations

import csv
import io
import zipfile
from pathlib import Path, PurePosixPath


WINDOWS_MAX_PATH_LENGTH = 259  # MAX_PATH minus the terminating null character.


class ZipExtractionPathError(ValueError):
    def __init__(self, output_folder: Path, longest_path: Path) -> None:
        self.output_folder = output_folder
        self.longest_path = longest_path
        super().__init__(
            "The extraction destination would create a path that Windows previews "
            f"and Microsoft Office cannot open ({_windows_path_length(longest_path)} "
            f"characters; maximum {WINDOWS_MAX_PATH_LENGTH}):\n{longest_path}\n\n"
            "Choose an extraction folder closer to the drive root. "
            "No files have been extracted from this ZIP."
        )


class ZipExtractionWarning(Exception):
    def __init__(self, output_folder: Path, unsafe_paths: list[str]) -> None:
        self.output_folder = output_folder
        self.unsafe_paths = unsafe_paths
        super().__init__(
            f"Skipped unsafe ZIP path(s); extracted safe files to {output_folder}:\n" + "\n".join(unsafe_paths)
        )


def short_name_for_sequence(sequence: int) -> str:
    if sequence < 1:
        raise ValueError("Sequence must be 1 or greater.")

    zero_based = sequence - 1
    group_index = zero_based // 9
    suffix = (zero_based % 9) + 1
    return f"{_letters_for_group(group_index)}{suffix}"


def _letters_for_group(group_index: int) -> str:
    if group_index < 26:
        return chr(ord("A") + group_index)

    return f"Z{_letters_for_group(group_index - 26)}"


def default_renamed_zip_path(source_zip: Path) -> Path:
    return source_zip.with_name(f"{source_zip.stem} - renamed.zip")


def default_extracted_folder_path(source_zip: Path, parent_folder: Path | None = None) -> Path:
    parent = Path(parent_folder) if parent_folder is not None else Path(source_zip).parent
    return parent / f"{Path(source_zip).stem} - renamed"


def unique_zip_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def unique_folder_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 2
    while True:
        candidate = path.with_name(f"{path.name} ({counter})")
        if not candidate.exists():
            return candidate
        counter += 1


def rename_zip_contents(source_zip: Path, output_zip: Path) -> Path:
    source_zip = Path(source_zip)
    output_zip = Path(output_zip)

    if source_zip.resolve() == output_zip.resolve():
        raise ValueError("Output ZIP path must be different from the original ZIP path.")

    mapping_rows: list[dict[str, str]] = []
    sequence = 0

    with zipfile.ZipFile(source_zip, "r") as source, zipfile.ZipFile(output_zip, "w") as target:
        for info in source.infolist():
            if info.is_dir():
                target.writestr(info, b"")
                continue

            sequence += 1
            original_path = PurePosixPath(info.filename)
            folder_path = original_path.parent.as_posix()
            if folder_path == ".":
                folder_path = ""
            new_name = f"{short_name_for_sequence(sequence)}{original_path.suffix}"
            new_path = new_name if not folder_path else f"{folder_path}/{new_name}"

            renamed_info = zipfile.ZipInfo(new_path, info.date_time)
            renamed_info.comment = info.comment
            renamed_info.extra = info.extra
            renamed_info.internal_attr = info.internal_attr
            renamed_info.external_attr = info.external_attr
            renamed_info.create_system = info.create_system
            renamed_info.compress_type = info.compress_type
            renamed_info._compresslevel = getattr(info, "_compresslevel", None)

            with source.open(info) as original_file:
                target.writestr(renamed_info, original_file.read())
            mapping_rows.append(
                {
                    "FolderPath": folder_path or "[ZIP ROOT]",
                    "GlobalSequence": str(sequence),
                    "NewName": new_name,
                    "NewPath": new_path,
                    "OriginalName": original_path.name,
                    "OriginalPath": info.filename,
                }
            )

        target.writestr("Name_Mapping.csv", _build_mapping_csv(mapping_rows))

    return output_zip


def rename_and_extract_zip_contents(source_zip: Path, output_folder: Path) -> Path:
    source_zip = Path(source_zip)
    output_folder = Path(output_folder)
    output_root = output_folder.resolve()

    mapping_rows: list[dict[str, str]] = []
    warnings: list[str] = []
    sequence = 0

    with zipfile.ZipFile(source_zip, "r") as source:
        _check_extraction_path_lengths(source, output_root)
        output_folder.mkdir(parents=True, exist_ok=False)
        for info in source.infolist():
            original_path = _safe_zip_path(info.filename)
            if original_path is None:
                warnings.append(info.filename)
                continue

            if info.is_dir():
                directory_path = _safe_output_path(output_root, original_path)
                directory_path.mkdir(parents=True, exist_ok=True)
                continue

            sequence += 1
            folder_path = original_path.parent.as_posix()
            if folder_path == ".":
                folder_path = ""
            new_name = f"{short_name_for_sequence(sequence)}{original_path.suffix}"
            new_path = new_name if not folder_path else f"{folder_path}/{new_name}"
            output_path = _unique_file_path(_safe_output_path(output_root, PurePosixPath(new_path)))
            output_path.parent.mkdir(parents=True, exist_ok=True)

            with source.open(info) as original_file, output_path.open("wb") as extracted_file:
                extracted_file.write(original_file.read())
            mapping_rows.append(
                {
                    "FolderPath": folder_path or "[ZIP ROOT]",
                    "GlobalSequence": str(sequence),
                    "NewName": output_path.name,
                    "NewPath": output_path.relative_to(output_root).as_posix(),
                    "OriginalName": original_path.name,
                    "OriginalPath": info.filename,
                }
            )

    mapping_path = _unique_file_path(output_root / "Name_Mapping.csv")
    mapping_path.write_text(_build_mapping_csv(mapping_rows), encoding="utf-8", newline="")

    if warnings:
        raise ZipExtractionWarning(output_folder, warnings)

    return output_folder


def _windows_path_length(path: Path) -> int:
    # Python can access extended paths, but Word and Explorer still use MAX_PATH.
    text = str(path.absolute())
    if text.startswith("\\\\?\\UNC\\"):
        text = "\\\\" + text[8:]
    elif text.startswith("\\\\?\\"):
        text = text[4:]
    return len(text.encode("utf-16-le")) // 2


def _check_extraction_path_lengths(source: zipfile.ZipFile, output_root: Path) -> None:
    longest_path = output_root / "Name_Mapping.csv"
    occupied_paths: set[Path] = set()
    sequence = 0
    for info in source.infolist():
        original_path = _safe_zip_path(info.filename)
        if original_path is None:
            continue
        if info.is_dir():
            relative_path = original_path
        else:
            sequence += 1
            relative_path = original_path.with_name(
                f"{short_name_for_sequence(sequence)}{original_path.suffix}"
            )
        target = _safe_output_path(output_root, relative_path)
        if not info.is_dir():
            target = _unique_file_path(target, occupied_paths)
        occupied_paths.add(target)
        occupied_paths.update(target.parents)
        if _windows_path_length(target) > _windows_path_length(longest_path):
            longest_path = target

    mapping_path = _unique_file_path(output_root / "Name_Mapping.csv", occupied_paths)
    if _windows_path_length(mapping_path) > _windows_path_length(longest_path):
        longest_path = mapping_path
    if _windows_path_length(longest_path) > WINDOWS_MAX_PATH_LENGTH:
        raise ZipExtractionPathError(output_root, longest_path)


def _safe_zip_path(filename: str) -> PurePosixPath | None:
    path = PurePosixPath(filename)
    if path.is_absolute() or any(part in ("", "..") for part in path.parts):
        return None
    return path


def _safe_output_path(output_root: Path, zip_path: PurePosixPath) -> Path:
    target = output_root.joinpath(*zip_path.parts)
    resolved = target.resolve(strict=False)
    if resolved != output_root and output_root not in resolved.parents:
        raise ValueError(f"Unsafe ZIP path skipped: {zip_path.as_posix()}")
    return target


def _unique_file_path(path: Path, occupied_paths: set[Path] | None = None) -> Path:
    def is_occupied(candidate: Path) -> bool:
        return candidate.exists() or (occupied_paths is not None and candidate in occupied_paths)

    if not is_occupied(path):
        return path

    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not is_occupied(candidate):
            return candidate
        counter += 1


def _build_mapping_csv(mapping_rows: list[dict[str, str]]) -> str:
    output = io.StringIO(newline="")
    fieldnames = ["RowType", "FolderPath", "GlobalSequence", "NewName", "NewPath", "OriginalName", "OriginalPath"]
    writer = csv.DictWriter(output, fieldnames=fieldnames)
    writer.writeheader()

    folders = list(dict.fromkeys(row["FolderPath"] for row in mapping_rows))
    rows_by_folder = {folder: [row for row in mapping_rows if row["FolderPath"] == folder] for folder in folders}
    for folder in folders:
        writer.writerow(
            {
                "RowType": "Folder",
                "FolderPath": folder,
                "GlobalSequence": "",
                "NewName": "",
                "NewPath": "",
                "OriginalName": "",
                "OriginalPath": "",
            }
        )
        for row in rows_by_folder[folder]:
            writer.writerow({"RowType": "File", **row})

    return output.getvalue()
