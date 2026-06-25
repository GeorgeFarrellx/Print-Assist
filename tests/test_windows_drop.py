from __future__ import annotations

import struct
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pythoncom

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


class _MultipleAttachmentDataObject:
    def __init__(
        self,
        descriptor_format: int,
        contents_format: int,
        descriptor_payload: bytes,
        attachment_payloads: list[bytes],
    ) -> None:
        self.descriptor_format = descriptor_format
        self.contents_format = contents_format
        self.descriptor_payload = descriptor_payload
        self.attachment_payloads = attachment_payloads

    def QueryGetData(self, format_etc: tuple[object, None, int, int, int]) -> None:
        clipboard_format, _, _, index, _ = format_etc
        if clipboard_format == self.descriptor_format and index == -1:
            return
        if clipboard_format == self.contents_format and 0 <= index < len(self.attachment_payloads):
            return
        raise RuntimeError("unsupported format")

    def GetData(self, format_etc: tuple[object, None, int, int, int]) -> SimpleNamespace:
        clipboard_format, _, _, index, _ = format_etc
        if clipboard_format == self.descriptor_format and index == -1:
            return SimpleNamespace(data=self.descriptor_payload)
        if clipboard_format == self.contents_format and 0 <= index < len(self.attachment_payloads):
            return SimpleNamespace(data=self.attachment_payloads[index])
        raise RuntimeError("unsupported format")


class _StorageDataObject:
    def __init__(self, contents_format: int, storage_tymed: int, storage: object) -> None:
        self.contents_format = contents_format
        self.storage_tymed = storage_tymed
        self.storage = storage

    def QueryGetData(self, format_etc: tuple[object, None, int, int, int]) -> None:
        clipboard_format, _, _, index, tymed = format_etc
        if (
            clipboard_format == self.contents_format
            and index == 0
            and tymed == self.storage_tymed
        ):
            return
        raise RuntimeError("unsupported format")

    def GetData(self, format_etc: tuple[object, None, int, int, int]) -> SimpleNamespace:
        self.QueryGetData(format_etc)
        return SimpleNamespace(data=self.storage)


class _GenericComInterface:
    def __init__(self, storage: object) -> None:
        self.storage = storage
        self.requested_iids: list[object] = []

    def QueryInterface(self, iid: object) -> object:
        self.requested_iids.append(iid)
        if iid == pythoncom.IID_IStorage:
            return self.storage
        raise RuntimeError("unsupported interface")


class _CombinedStorageDataObject:
    def __init__(
        self,
        contents_format: int,
        combined_tymed: int,
        storage_tymed: int,
        storage: object,
    ) -> None:
        self.contents_format = contents_format
        self.combined_tymed = combined_tymed
        self.storage_tymed = storage_tymed
        self.generic_interface = _GenericComInterface(storage)

    def QueryGetData(self, format_etc: tuple[object, None, int, int, int]) -> None:
        clipboard_format, _, _, index, tymed = format_etc
        if (
            clipboard_format == self.contents_format
            and index == 0
            and tymed == self.combined_tymed
        ):
            return
        raise RuntimeError("request all supported media together")

    def GetData(self, format_etc: tuple[object, None, int, int, int]) -> SimpleNamespace:
        self.QueryGetData(format_etc)
        return SimpleNamespace(data=self.generic_interface, tymed=self.storage_tymed)


def _unicode_descriptor_payload(names: list[str]) -> bytes:
    descriptors = []
    for name in names:
        descriptor = bytearray(592)
        encoded_name = name.encode("utf-16le") + b"\x00\x00"
        descriptor[-520 : -520 + len(encoded_name)] = encoded_name
        descriptors.append(descriptor)
    return struct.pack("<I", len(descriptors)) + b"".join(descriptors)


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

    def test_two_outlook_jpg_attachments_are_materialised_in_one_drop(self) -> None:
        descriptor_format = 1001
        contents_format = 1003
        names = ["letter 1.jpg", "letter 2.jpg"]
        payloads = [b"\xff\xd8first jpg\xff\xd9", b"\xff\xd8second jpg\xff\xd9"]
        data_object = _MultipleAttachmentDataObject(
            descriptor_format,
            contents_format,
            _unicode_descriptor_payload(names),
            payloads,
        )

        with (
            tempfile.TemporaryDirectory() as temp_dir,
            patch("print_assist.windows_drop.WINDOWS_CF_HDROP", 15),
            patch("print_assist.windows_drop.WINDOWS_FILEGROUPDESCRIPTORW", descriptor_format),
            patch("print_assist.windows_drop.WINDOWS_FILEGROUPDESCRIPTORA", 1002),
            patch("print_assist.windows_drop.WINDOWS_FILECONTENTS", contents_format),
            patch("print_assist.windows_drop.WINDOWS_TYMED_HGLOBAL", 1),
            patch("print_assist.windows_drop.WINDOWS_TYMED_ISTREAM", 4),
        ):
            fake_app = SimpleNamespace(_get_outlook_drop_temp_dir=lambda: Path(temp_dir))
            received_paths: list[str] = []
            errors: list[str] = []
            target = NativeWindowsDropTarget(
                on_paths=received_paths.extend,
                materialise_virtual_files=lambda attachment_names, attachment_payloads: (
                    PrintAssistApp._materialise_outlook_virtual_attachments(
                        fake_app,
                        attachment_names,
                        attachment_payloads,
                    )
                ),
                on_error=errors.append,
            )

            effect = target.Drop(data_object, 0, (0, 0), 1)

            self.assertEqual(effect, 1)
            self.assertEqual(errors, [])
            self.assertEqual([Path(path).name for path in received_paths], names)
            self.assertEqual([Path(path).read_bytes() for path in received_paths], payloads)

    def test_outlook_msg_istorage_payload_is_serialised_to_compound_file_bytes(self) -> None:
        contents_format = 1003
        storage_tymed = 8
        lock_bytes = pythoncom.CreateILockBytesOnHGlobal()
        source_storage = pythoncom.StgCreateDocfileOnILockBytes(lock_bytes, 0x1012, 0)
        stream = source_storage.CreateStream("__properties_version1.0", 0x1012, 0, 0)
        stream.Write(b"message properties")
        stream = None
        source_storage.Commit(0)
        data_object = _StorageDataObject(contents_format, storage_tymed, source_storage)

        with (
            patch("print_assist.windows_drop.WINDOWS_FILECONTENTS", contents_format),
            patch("print_assist.windows_drop.WINDOWS_TYMED_ISTREAM", 4),
            patch("print_assist.windows_drop.WINDOWS_TYMED_ISTORAGE", storage_tymed),
            patch("print_assist.windows_drop.WINDOWS_TYMED_HGLOBAL", 1),
        ):
            target = NativeWindowsDropTarget(
                on_paths=lambda paths: None,
                materialise_virtual_files=lambda names, payloads: [],
                on_error=lambda details: None,
            )
            payload = target._read_virtual_file_bytes(data_object, 0)

        self.assertIsNotNone(payload)
        self.assertTrue(payload.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"))

    def test_outlook_msg_accepts_combined_tymed_and_generic_istorage_interface(self) -> None:
        contents_format = 1003
        stream_tymed = 4
        storage_tymed = 8
        hglobal_tymed = 1
        combined_tymed = stream_tymed | storage_tymed | hglobal_tymed
        lock_bytes = pythoncom.CreateILockBytesOnHGlobal()
        source_storage = pythoncom.StgCreateDocfileOnILockBytes(lock_bytes, 0x1012, 0)
        stream = source_storage.CreateStream("__properties_version1.0", 0x1012, 0, 0)
        stream.Write(b"message properties")
        stream = None
        source_storage.Commit(0)
        data_object = _CombinedStorageDataObject(
            contents_format,
            combined_tymed,
            storage_tymed,
            source_storage,
        )

        with (
            patch("print_assist.windows_drop.WINDOWS_FILECONTENTS", contents_format),
            patch("print_assist.windows_drop.WINDOWS_TYMED_ISTREAM", stream_tymed),
            patch("print_assist.windows_drop.WINDOWS_TYMED_ISTORAGE", storage_tymed),
            patch("print_assist.windows_drop.WINDOWS_TYMED_HGLOBAL", hglobal_tymed),
            patch("print_assist.windows_drop.WINDOWS_IID_ISTORAGE", pythoncom.IID_IStorage),
        ):
            target = NativeWindowsDropTarget(
                on_paths=lambda paths: None,
                materialise_virtual_files=lambda names, payloads: [],
                on_error=lambda details: None,
            )
            payload = target._read_virtual_file_bytes(data_object, 0)

        self.assertIsNotNone(payload)
        self.assertTrue(payload.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"))
        self.assertEqual(
            data_object.generic_interface.requested_iids,
            [pythoncom.IID_IStorage],
        )

    def test_drop_handler_exception_is_contained_at_com_boundary(self) -> None:
        errors: list[str] = []
        target = NativeWindowsDropTarget(
            on_paths=lambda paths: (_ for _ in ()).throw(RuntimeError("UI callback failed")),
            materialise_virtual_files=lambda names, payloads: [],
            on_error=errors.append,
        )
        target._extract_paths = lambda data_object: ["attachment.jpg"]  # type: ignore[method-assign]

        effect = target.Drop(object(), 0, (0, 0), 1)

        self.assertEqual(effect, 0)
        self.assertEqual(errors, ["UI callback failed"])


if __name__ == "__main__":
    unittest.main()
