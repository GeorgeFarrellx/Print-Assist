from __future__ import annotations

import os
import struct
from collections.abc import Callable, Iterable

if os.name == "nt":
    try:
        import pythoncom
        import win32clipboard
        import win32con
        from win32com.server.util import wrap as win32com_wrap
        from win32com.shell import shell
    except Exception:
        pythoncom = None
        win32clipboard = None
        win32con = None
        win32com_wrap = None
        shell = None
else:
    pythoncom = None
    win32clipboard = None
    win32con = None
    win32com_wrap = None
    shell = None


WINDOWS_NATIVE_DROP_AVAILABLE = bool(
    os.name == "nt"
    and pythoncom is not None
    and win32clipboard is not None
    and win32con is not None
    and win32com_wrap is not None
    and shell is not None
)

WINDOWS_FILEGROUPDESCRIPTORW = (
    win32clipboard.RegisterClipboardFormat("FileGroupDescriptorW")
    if WINDOWS_NATIVE_DROP_AVAILABLE
    else None
)
WINDOWS_FILEGROUPDESCRIPTORA = (
    win32clipboard.RegisterClipboardFormat("FileGroupDescriptor")
    if WINDOWS_NATIVE_DROP_AVAILABLE
    else None
)
WINDOWS_FILECONTENTS = (
    win32clipboard.RegisterClipboardFormat("FileContents")
    if WINDOWS_NATIVE_DROP_AVAILABLE
    else None
)
WINDOWS_DROPEFFECT_COPY = getattr(win32con, "DROPEFFECT_COPY", 1) if win32con else 1
WINDOWS_DROPEFFECT_NONE = getattr(win32con, "DROPEFFECT_NONE", 0) if win32con else 0
WINDOWS_CF_HDROP = getattr(win32con, "CF_HDROP", 15) if win32con else 15
WINDOWS_DVASPECT_CONTENT = getattr(pythoncom, "DVASPECT_CONTENT", 1) if pythoncom else 1
WINDOWS_TYMED_HGLOBAL = getattr(pythoncom, "TYMED_HGLOBAL", 1) if pythoncom else 1
WINDOWS_TYMED_ISTREAM = getattr(pythoncom, "TYMED_ISTREAM", 4) if pythoncom else 4
WINDOWS_TYMED_ISTORAGE = getattr(pythoncom, "TYMED_ISTORAGE", 8) if pythoncom else 8
WINDOWS_IID_ISTREAM = getattr(pythoncom, "IID_IStream", None) if pythoncom else None
WINDOWS_IID_ISTORAGE = getattr(pythoncom, "IID_IStorage", None) if pythoncom else None
WINDOWS_STGM_CREATE_READWRITE_EXCLUSIVE = 0x1000 | 0x0002 | 0x0010
WINDOWS_STATFLAG_NONAME = 1
WINDOWS_STGC_DEFAULT = 0


PathHandler = Callable[[list[str]], None]
VirtualFileMaterialiser = Callable[[list[str], Iterable[bytes | None]], list[str]]
ErrorHandler = Callable[[str], None]


def ole_initialize() -> bool:
    if not WINDOWS_NATIVE_DROP_AVAILABLE:
        return False
    try:
        pythoncom.OleInitialize()
        return True
    except Exception:
        return False


def ole_uninitialize() -> None:
    if pythoncom is None:
        return
    try:
        pythoncom.OleUninitialize()
    except Exception:
        pass


def register_drop_target(hwnd: int, target: NativeWindowsDropTarget) -> object | None:
    if not WINDOWS_NATIVE_DROP_AVAILABLE:
        return None
    wrapped = win32com_wrap(target, iid=pythoncom.IID_IDropTarget, useDispatcher=0)
    pythoncom.RegisterDragDrop(hwnd, wrapped)
    return wrapped


def revoke_drop_target(hwnd: int) -> None:
    if not WINDOWS_NATIVE_DROP_AVAILABLE:
        return
    try:
        pythoncom.RevokeDragDrop(hwnd)
    except Exception:
        pass


class NativeWindowsDropTarget:
    """OLE drop target for Explorer files and Outlook virtual attachments."""

    _com_interfaces_ = [pythoncom.IID_IDropTarget] if pythoncom else []
    _public_methods_ = ["DragEnter", "DragOver", "DragLeave", "Drop"]

    def __init__(
        self,
        on_paths: PathHandler,
        materialise_virtual_files: VirtualFileMaterialiser,
        on_error: ErrorHandler,
    ) -> None:
        self._on_paths = on_paths
        self._materialise_virtual_files = materialise_virtual_files
        self._on_error = on_error
        self._supports_current_drag = False

    @staticmethod
    def _format_etc(clipboard_format: int | None, tymed: int, index: int = -1) -> tuple[object, None, int, int, int]:
        return (
            clipboard_format,
            None,
            WINDOWS_DVASPECT_CONTENT,
            index,
            tymed,
        )

    def _query_get_data(self, data_object: object, clipboard_format: int | None, tymed: int, index: int = -1) -> bool:
        if data_object is None or clipboard_format is None:
            return False
        try:
            data_object.QueryGetData(self._format_etc(clipboard_format, tymed, index=index))
            return True
        except Exception:
            return False

    def _supports_drag(self, data_object: object) -> bool:
        return any(
            (
                self._query_get_data(data_object, WINDOWS_CF_HDROP, WINDOWS_TYMED_HGLOBAL),
                self._query_get_data(data_object, WINDOWS_FILEGROUPDESCRIPTORW, WINDOWS_TYMED_HGLOBAL),
                self._query_get_data(data_object, WINDOWS_FILEGROUPDESCRIPTORA, WINDOWS_TYMED_HGLOBAL),
            )
        )

    def _extract_hdrop_paths(self, data_object: object) -> list[str]:
        if not self._query_get_data(data_object, WINDOWS_CF_HDROP, WINDOWS_TYMED_HGLOBAL):
            return []
        try:
            medium = data_object.GetData(self._format_etc(WINDOWS_CF_HDROP, WINDOWS_TYMED_HGLOBAL))
            count = shell.DragQueryFile(medium.data_handle, -1)
            return [shell.DragQueryFile(medium.data_handle, index) for index in range(count)]
        except Exception:
            return []

    def _extract_virtual_file_names(self, data_object: object) -> list[str]:
        # FILEGROUPDESCRIPTOR starts with a count followed by fixed-size
        # FILEDESCRIPTOR records. cFileName is the final field in each record.
        for clipboard_format, descriptor_size, name_size, encoding in (
            (WINDOWS_FILEGROUPDESCRIPTORW, 592, 520, "utf-16le"),
            (WINDOWS_FILEGROUPDESCRIPTORA, 332, 260, "mbcs"),
        ):
            if not self._query_get_data(data_object, clipboard_format, WINDOWS_TYMED_HGLOBAL):
                continue
            try:
                medium = data_object.GetData(self._format_etc(clipboard_format, WINDOWS_TYMED_HGLOBAL))
                raw = bytes(medium.data or b"")
                if len(raw) < 4:
                    continue
                count = struct.unpack_from("<I", raw, 0)[0]
                names = []
                for index in range(count):
                    start = 4 + (index * descriptor_size)
                    end = start + descriptor_size
                    if end > len(raw):
                        break
                    name_bytes = raw[start:end][-name_size:]
                    try:
                        name = name_bytes.decode(encoding, errors="ignore")
                    except LookupError:
                        name = name_bytes.decode("latin-1", errors="ignore")
                    names.append(name.split("\x00", 1)[0].strip())
                if names:
                    return names
            except Exception:
                continue
        return []

    @staticmethod
    def _as_com_interface(value: object, iid: object | None) -> object:
        if value is None or iid is None or not hasattr(value, "QueryInterface"):
            return value
        try:
            return value.QueryInterface(iid)
        except Exception:
            return value

    def _read_stream_bytes(self, value: object) -> bytes | None:
        try:
            stream = self._as_com_interface(value, WINDOWS_IID_ISTREAM)
            stream.Seek(0, 0)
            chunks = []
            while True:
                chunk = stream.Read(65536)
                if not chunk:
                    break
                chunks.append(chunk)
            return b"".join(chunks)
        except Exception:
            return None

    def _read_storage_bytes(self, value: object) -> bytes | None:
        # Whole Outlook messages are commonly supplied as an IStorage compound
        # file. Real Outlook can expose the STGMEDIUM value as IUnknown, so
        # explicitly query it for IStorage before copying it to a docfile.
        source_storage = None
        destination_storage = None
        lock_bytes = None
        try:
            source_storage = self._as_com_interface(value, WINDOWS_IID_ISTORAGE)
            lock_bytes = pythoncom.CreateILockBytesOnHGlobal()
            destination_storage = pythoncom.StgCreateDocfileOnILockBytes(
                lock_bytes,
                WINDOWS_STGM_CREATE_READWRITE_EXCLUSIVE,
                0,
            )
            source_storage.CopyTo([], None, destination_storage)
            destination_storage.Commit(WINDOWS_STGC_DEFAULT)
            size = int(lock_bytes.Stat(WINDOWS_STATFLAG_NONAME)[2])
            return bytes(lock_bytes.ReadAt(0, size))
        except Exception:
            return None
        finally:
            source_storage = None
            destination_storage = None
            lock_bytes = None

    def _read_medium_bytes(self, medium: object, requested_tymed: int) -> bytes | None:
        actual_tymed = int(getattr(medium, "tymed", requested_tymed))
        value = getattr(medium, "data", None)

        if actual_tymed & WINDOWS_TYMED_ISTREAM:
            payload = self._read_stream_bytes(value)
            if payload is not None:
                return payload
        if actual_tymed & WINDOWS_TYMED_ISTORAGE:
            payload = self._read_storage_bytes(value)
            if payload is not None:
                return payload
        if actual_tymed & WINDOWS_TYMED_HGLOBAL:
            try:
                return bytes(value or b"")
            except Exception:
                pass
        return None

    def _read_virtual_file_bytes(self, data_object: object, index: int) -> bytes | None:
        # FORMATETC.tymed is a bitmask. Asking for all supported media together
        # lets Outlook choose the representation it can render for this item.
        # Keep individual requests as fallbacks for stricter data providers.
        combined_tymed = (
            WINDOWS_TYMED_ISTREAM
            | WINDOWS_TYMED_ISTORAGE
            | WINDOWS_TYMED_HGLOBAL
        )
        attempted: set[int] = set()
        for requested_tymed in (
            combined_tymed,
            WINDOWS_TYMED_ISTREAM,
            WINDOWS_TYMED_ISTORAGE,
            WINDOWS_TYMED_HGLOBAL,
        ):
            if requested_tymed in attempted:
                continue
            attempted.add(requested_tymed)
            try:
                medium = data_object.GetData(
                    self._format_etc(
                        WINDOWS_FILECONTENTS,
                        requested_tymed,
                        index=index,
                    )
                )
            except Exception:
                continue
            payload = self._read_medium_bytes(medium, requested_tymed)
            if payload is not None:
                return payload
        return None

    def _extract_paths(self, data_object: object) -> list[str]:
        paths = self._extract_hdrop_paths(data_object)
        if paths:
            return paths

        names = self._extract_virtual_file_names(data_object)
        if not names:
            return []
        payloads = (self._read_virtual_file_bytes(data_object, index) for index in range(len(names)))
        paths = self._materialise_virtual_files(names, payloads)
        if not paths:
            raise RuntimeError(
                "Outlook supplied the item name but not a readable file payload "
                "(tried IStream, IStorage, and HGLOBAL formats)."
            )
        return paths

    def DragEnter(self, data_object: object, key_state: int, point: object, effect: int) -> int:
        self._supports_current_drag = self._supports_drag(data_object)
        return WINDOWS_DROPEFFECT_COPY if self._supports_current_drag else WINDOWS_DROPEFFECT_NONE

    def DragOver(self, key_state: int, point: object, effect: int) -> int:
        return WINDOWS_DROPEFFECT_COPY if self._supports_current_drag else WINDOWS_DROPEFFECT_NONE

    def DragLeave(self) -> None:
        self._supports_current_drag = False

    def Drop(self, data_object: object, key_state: int, point: object, effect: int) -> int:
        self._supports_current_drag = False
        try:
            paths = self._extract_paths(data_object)
            self._on_paths(paths)
        except Exception as exc:
            try:
                self._on_error(str(exc))
            except Exception:
                pass
            return WINDOWS_DROPEFFECT_NONE

        return WINDOWS_DROPEFFECT_COPY if paths else WINDOWS_DROPEFFECT_NONE
