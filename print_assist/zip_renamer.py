from __future__ import annotations

import csv
import io
import zipfile
from pathlib import Path, PurePosixPath


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


def unique_zip_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
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
