from __future__ import annotations

import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path


SHORTCUT_INSTALLERS = ("Create Print Assist Shortcut.bat", "Create Print Assist Shortcut.ps1")
BUILD_INFO_NAME = "print-assist-build.json"


def app_source_hashes(project_root: Path) -> dict[str, str]:
    sources = [project_root / "main.py", project_root / "requirements.txt"]
    sources.extend((project_root / "print_assist").rglob("*.py"))
    sources.extend(path for path in (project_root / "print_assist" / "assets").rglob("*") if path.is_file())
    hashes = {}
    for path in sorted(sources):
        if not path.is_file():
            continue
        content = path.read_bytes()
        if path.suffix.lower() in {".py", ".txt"}:
            content = content.replace(b"\r\n", b"\n")
        hashes[path.relative_to(project_root).as_posix()] = hashlib.sha256(content).hexdigest()
    return hashes


def finish_windows_build(project_root: Path, output_dir: Path) -> Path:
    """Record the exact app sources and ship the shortcut installer with the bundle."""
    executable = output_dir / "PrintAssist.exe"
    metadata = {
        "schema_version": 1,
        "built_at_utc": datetime.now(timezone.utc).isoformat(),
        "executable_sha256": hashlib.sha256(executable.read_bytes()).hexdigest(),
        "source_hashes": app_source_hashes(project_root),
    }
    info_path = output_dir / BUILD_INFO_NAME
    info_path.write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    for name in SHORTCUT_INSTALLERS:
        shutil.copy2(project_root / name, output_dir / name)
    return info_path
