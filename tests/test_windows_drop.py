from __future__ import annotations

import struct
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from print_assist.app import PrintAssistApp
from print_assist.windows_drop import NativeWindowsDropTarget


class _DescriptorDataObject:
    def __init__(self, clipboard_format: int, payload: bytes) -> None:
        self.clipboard_format = clipboard_format
        self.payload = payload

    def QueryGetData(self, format_etc: tuple[object, None, int, int, int]) -> None:
        if format_etc[0] != self.clipboard_format:
            raise RuntimeError("unsupported format")

    def GetData(self, format_etc: tuple[object, None, int, int, int]) -> SimpleNamespace:
        if format_etc[0] != self.clipboard_format:
            raise RuntimeError("unsupported format")
        return SimpleNamespace(data=self.payload)


class WindowsDropTests(unittest.TestCase):
    def test_outlook_unicode_descriptor_filename_is_decoded(self) -> None:
        clipboard_format = 1001
        filename = "Statement 12-MAR-25.pdf"
        descriptor = bytearray(592)
        encoded_name = filename.encode("utf-16le") + b"\x00\x00"
        descriptor[-520 : -520 + len(encoded_name)] = encoded_name
        payload = struct.pack("<I", 1) + descriptor

        with (
            patch("print_assist.windows_drop.WINDOWS_FILEGROUPDESCRIPTORW", clipboard_format),
            patch("print_assist.windows_drop.WINDOWS_FILEGROUPDESCRIPTORA", 1002),
            patch("print_assist.windows_drop.WINDOWS_TYMED_HGLOBAL", 1),
        ):
            target = NativeWindowsDropTarget(
                on_paths=lambda paths: None,
                materialise_virtual_files=lambda names, payloads: [],
                on_error=lambda details: None,
            )
            names = target._extract_virtual_file_names(_DescriptorDataObject(clipboard_format, payload))

        self.assertEqual(names, [filename])

    def test_outlook_pdf_payload_is_materialised_with_pdf_extension(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            fake_app = SimpleNamespace(_get_outlook_drop_temp_dir=lambda: Path(temp_dir))
            payload = b"%PDF-1.7\nstatement bytes"

            paths = PrintAssistApp._materialise_outlook_virtual_attachments(
                fake_app,
                ["Statement without extension"],
                [payload],
            )

            self.assertEqual(len(paths), 1)
            output_path = Path(paths[0])
            self.assertEqual(output_path.suffix, ".pdf")
            self.assertEqual(output_path.read_bytes(), payload)


if __name__ == "__main__":
    unittest.main()
