from __future__ import annotations

import threading
import time
from pathlib import Path

WORD_EXTENSIONS = {".doc", ".docx"}
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
MSG_EXTENSIONS = {".msg"}


PDF_FORMAT = 17
WORD_ALERTS_NONE = 0
EXCEL_XLTYPE_PDF = 0
EXCEL_UPDATE_LINKS_NEVER = 0
EXCEL_CORRUPT_LOAD_NORMAL = 0
OUTLOOK_OLMSG_UNICODE = 9
OUTLOOK_OL_MHTML = 10
OUTLOOK_PDF_PRINTER_NAME = "Microsoft Print to PDF"
OUTLOOK_PDF_DIALOG_TITLES = {"Save Print Output As", "Save As"}
OUTLOOK_PDF_DIALOG_TIMEOUT_SECONDS = 30
OUTLOOK_PDF_OUTPUT_TIMEOUT_SECONDS = 15

_OUTLOOK_PRINT_LOCK = threading.Lock()


def _require_windows() -> None:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception as exc:
        raise RuntimeError("pywin32/COM automation is unavailable on this system") from exc


def _word_to_pdf(source_path: Path, output_path: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = WORD_ALERTS_NONE
        doc = word.Documents.Open(str(source_path), ReadOnly=True, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(str(output_path), PDF_FORMAT)
    finally:
        if doc is not None:
            doc.Close(False)
            doc = None
        if word is not None:
            word.Quit()
            word = None
        pythoncom.CoUninitialize()


def _excel_to_pdf(source_path: Path, output_path: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(
            str(source_path),
            UpdateLinks=EXCEL_UPDATE_LINKS_NEVER,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            CorruptLoad=EXCEL_CORRUPT_LOAD_NORMAL,
        )
        workbook.ExportAsFixedFormat(EXCEL_XLTYPE_PDF, str(output_path))
    finally:
        if workbook is not None:
            workbook.Close(False)
            workbook = None
        if excel is not None:
            excel.Quit()
            excel = None
        pythoncom.CoUninitialize()


def _visible_top_level_windows(win32gui: object) -> set[int]:
    handles: set[int] = set()

    def collect(hwnd: int, extra: object) -> bool:
        _ = extra
        if win32gui.IsWindowVisible(hwnd):
            handles.add(hwnd)
        return True

    win32gui.EnumWindows(collect, None)
    return handles


def _child_windows(win32gui: object, parent_hwnd: int) -> list[int]:
    handles: list[int] = []

    def collect(hwnd: int, extra: object) -> bool:
        _ = extra
        handles.append(hwnd)
        return True

    win32gui.EnumChildWindows(parent_hwnd, collect, None)
    return handles


def _handle_outlook_pdf_save_dialog(
    output_path: Path,
    ignored_windows: set[int],
    result: dict[str, object],
) -> None:
    import win32con
    import win32gui

    deadline = time.monotonic() + OUTLOOK_PDF_DIALOG_TIMEOUT_SECONDS
    matching_dialogs: set[int] = set()

    while time.monotonic() < deadline:
        for hwnd in _visible_top_level_windows(win32gui) - ignored_windows:
            title = win32gui.GetWindowText(hwnd).strip()
            if title not in OUTLOOK_PDF_DIALOG_TITLES:
                continue
            matching_dialogs.add(hwnd)

            children = _child_windows(win32gui, hwnd)
            edits = [
                child
                for child in children
                if win32gui.IsWindowVisible(child) and win32gui.GetClassName(child) == "Edit"
            ]
            buttons = [
                child
                for child in children
                if win32gui.IsWindowVisible(child) and win32gui.GetClassName(child) == "Button"
            ]
            if not edits:
                continue

            win32gui.SendMessage(edits[0], win32con.WM_SETTEXT, 0, str(output_path))
            save_buttons = [
                button
                for button in buttons
                if win32gui.GetWindowText(button).replace("&", "").strip().lower() == "save"
            ]
            if save_buttons:
                win32gui.SendMessage(save_buttons[0], win32con.BM_CLICK, 0, 0)
            else:
                win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
            result["handled"] = True
            return
        time.sleep(0.1)

    for hwnd in matching_dialogs:
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass
    result["error"] = "The Microsoft Print to PDF save dialog could not be completed."


def _wait_for_pdf_output(output_path: Path) -> None:
    deadline = time.monotonic() + OUTLOOK_PDF_OUTPUT_TIMEOUT_SECONDS
    previous_size = -1
    stable_checks = 0

    while time.monotonic() < deadline:
        if output_path.exists():
            current_size = output_path.stat().st_size
            if current_size > 0 and current_size == previous_size:
                stable_checks += 1
                if stable_checks >= 2:
                    return
            else:
                stable_checks = 0
            previous_size = current_size
        time.sleep(0.2)

    raise RuntimeError("Microsoft Print to PDF did not create the Outlook message PDF.")


def _msg_to_pdf(source_path: Path, output_path: Path, temp_dir: Path) -> None:
    _ = temp_dir
    import pythoncom
    import win32com.client
    import win32gui
    import win32print

    with _OUTLOOK_PRINT_LOCK:
        pythoncom.CoInitialize()
        outlook = None
        namespace = None
        message = None
        original_printer = None
        dialog_thread = None
        dialog_result: dict[str, object] = {}
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            if output_path.exists():
                output_path.unlink()

            available_printers = {
                printer[2]
                for printer in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
                )
            }
            if OUTLOOK_PDF_PRINTER_NAME not in available_printers:
                raise RuntimeError(f"'{OUTLOOK_PDF_PRINTER_NAME}' is not installed.")

            original_printer = win32print.GetDefaultPrinter()
            ignored_windows = _visible_top_level_windows(win32gui)
            win32print.SetDefaultPrinter(OUTLOOK_PDF_PRINTER_NAME)
            if win32print.GetDefaultPrinter() != OUTLOOK_PDF_PRINTER_NAME:
                raise RuntimeError("Windows did not select Microsoft Print to PDF.")

            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            message = namespace.OpenSharedItem(str(source_path))

            dialog_thread = threading.Thread(
                target=_handle_outlook_pdf_save_dialog,
                args=(output_path, ignored_windows, dialog_result),
                daemon=True,
            )
            dialog_thread.start()
            message.PrintOut()
            dialog_thread.join(timeout=OUTLOOK_PDF_DIALOG_TIMEOUT_SECONDS + 2)

            if dialog_thread.is_alive():
                raise RuntimeError("Timed out while closing the Microsoft Print to PDF dialog.")
            if dialog_result.get("error"):
                raise RuntimeError(str(dialog_result["error"]))
            if not dialog_result.get("handled"):
                raise RuntimeError("The Microsoft Print to PDF save dialog did not appear.")

            _wait_for_pdf_output(output_path)
        except Exception as exc:
            raise RuntimeError(
                f"Failed to print Outlook message in Memo Style: {source_path.name}"
            ) from exc
        finally:
            message = None
            namespace = None
            outlook = None
            if original_printer is not None:
                try:
                    win32print.SetDefaultPrinter(original_printer)
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def convert_to_pdf_if_needed(source_path: Path, temp_dir: Path) -> Path:
    suffix = source_path.suffix.lower()
    if suffix == ".pdf":
        return source_path

    _require_windows()
    output_path = temp_dir / f"{source_path.stem}.converted.pdf"

    if suffix in WORD_EXTENSIONS:
        _word_to_pdf(source_path, output_path)
        return output_path
    if suffix in EXCEL_EXTENSIONS:
        _excel_to_pdf(source_path, output_path)
        return output_path
    if suffix in MSG_EXTENSIONS:
        _msg_to_pdf(source_path, output_path, temp_dir)
        return output_path

    raise ValueError(f"Unsupported Office/MSG file extension: {source_path.suffix}")
