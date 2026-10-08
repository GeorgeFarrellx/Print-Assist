from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from build_support import app_source_hashes, finish_windows_build


PROJECT_ROOT = Path(__file__).resolve().parents[1]
INSTALLERS = ("Create Print Assist Shortcut.bat", "Create Print Assist Shortcut.ps1")


def create_project(root: Path) -> None:
    root.mkdir(parents=True)
    (root / "main.py").write_text("print('current app')\n", encoding="utf-8")
    (root / "requirements.txt").write_text("Pillow\n", encoding="utf-8")
    package = root / "print_assist"
    package.mkdir()
    (package / "app.py").write_text("CURRENT = True\n", encoding="utf-8")
    assets = package / "assets"
    assets.mkdir()
    (assets / "sample.png").write_bytes(b"binary fixture\r\n")
    for name in INSTALLERS:
        shutil.copy2(PROJECT_ROOT / name, root / name)


def create_build(root: Path) -> Path:
    output = root / "dist" / "PrintAssist"
    output.mkdir(parents=True)
    (output / "PrintAssist.exe").write_bytes(b"Current executable fixture")
    finish_windows_build(root, output)
    return output


class BuildSupportTests(unittest.TestCase):
    def test_build_record_binds_the_executable_to_sources_and_ships_installers(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            create_project(root)
            output = create_build(root)
            metadata = json.loads((output / "print-assist-build.json").read_text(encoding="utf-8"))
            self.assertEqual(metadata["source_hashes"], app_source_hashes(root))
            self.assertEqual(metadata["executable_sha256"], hashlib.sha256((output / "PrintAssist.exe").read_bytes()).hexdigest())
            self.assertEqual(metadata["schema_version"], 1)
            for name in INSTALLERS:
                self.assertEqual((output / name).read_bytes(), (root / name).read_bytes())

    def test_text_line_endings_do_not_make_an_identical_app_look_stale(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "project"
            create_project(root)
            (root / "main.py").write_bytes(b"print('current app')\n")
            hashes = app_source_hashes(root)
            (root / "main.py").write_bytes(b"print('current app')\r\n")
            self.assertEqual(app_source_hashes(root), hashes)


@unittest.skipUnless(sys.platform == "win32", "Windows shortcut integration")
class ShortcutInstallerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name) / "App & client's files"
        create_project(self.root)
        self.desktop = Path(self.temporary.name) / "Desktop"

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def _install(self, root: Path | None = None) -> subprocess.CompletedProcess[str]:
        directory = root or self.root
        env = dict(os.environ, PRINT_ASSIST_NO_PAUSE="1", PRINT_ASSIST_SHORTCUT_DIR=str(self.desktop))
        # cmd /c needs its outer quotes when a batch filename contains spaces.
        command = f'cmd.exe /d /s /c ""{directory / INSTALLERS[0]}""'
        return subprocess.run(command, env=env, capture_output=True, text=True, timeout=20)

    def _shortcut(self) -> dict[str, str]:
        env = dict(os.environ, PRINT_ASSIST_SHORTCUT_DIR=str(self.desktop))
        command = "$shell=New-Object -ComObject WScript.Shell; $link=$shell.CreateShortcut((Join-Path $env:PRINT_ASSIST_SHORTCUT_DIR 'Print Assist.lnk')); @{Target=$link.TargetPath; Directory=$link.WorkingDirectory; Icon=$link.IconLocation} | ConvertTo-Json -Compress"
        result = subprocess.run(["powershell.exe", "-NoProfile", "-Command", command], env=env, capture_output=True, text=True, check=True, timeout=20)
        return json.loads(result.stdout)

    def test_current_build_wins_over_legacy_copies_and_preserves_source_path_quoting(self) -> None:
        output = create_build(self.root)
        (self.root / "PrintAssist.exe").write_bytes(b"Old root copy")
        (self.root / "dist" / "PrintAssist.exe").write_bytes(b"Old single-file copy")
        (self.root / "main.py").write_bytes(b"print('current app')\r\n")
        result = self._install()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        shortcut = self._shortcut()
        self.assertEqual(shortcut["Target"], str(output / "PrintAssist.exe"))
        self.assertEqual(shortcut["Directory"], str(output))

    def test_source_changes_do_not_replace_the_existing_shortcut(self) -> None:
        create_build(self.root)
        result = self._install()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        original = (self.desktop / "Print Assist.lnk").read_bytes()
        (self.root / "print_assist" / "app.py").write_text("CURRENT = 'updated'\n", encoding="utf-8")
        result = self._install()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Rebuild first", result.stderr)
        self.assertEqual((self.desktop / "Print Assist.lnk").read_bytes(), original)

    def test_added_modules_and_replaced_executables_are_rejected(self) -> None:
        output = create_build(self.root)
        added = self.root / "print_assist" / "new_feature.py"
        added.write_text("NEW = True\n", encoding="utf-8")
        result = self._install()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Rebuild first", result.stderr)
        added.unlink()
        (output / "PrintAssist.exe").write_bytes(b"Replaced with old executable")
        result = self._install()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Rebuild first", result.stderr)
        self.assertFalse((self.desktop / "Print Assist.lnk").exists())

    def test_distributed_app_folder_can_create_its_own_shortcut(self) -> None:
        output = create_build(self.root)
        result = self._install(output)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(self._shortcut()["Target"], str(output / "PrintAssist.exe"))

    def test_missing_executable_leaves_no_shortcut(self) -> None:
        result = self._install()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("executable was not found", result.stderr)
        self.assertFalse((self.desktop / "Print Assist.lnk").exists())


if __name__ == "__main__":
    unittest.main()
