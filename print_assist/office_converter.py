from __future__ import annotations

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
        if word is not None:
            word.Quit()
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
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def _msg_to_pdf(source_path: Path, output_path: Path, temp_dir: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    outlook = None
    message = None
    temp_mht = temp_dir / f"{source_path.stem}_msg_export.mht"
    try:
        outlook = win32com.client.DispatchEx("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        message = namespace.OpenSharedItem(str(source_path))
        message.SaveAs(str(temp_mht), OUTLOOK_OL_MHTML)
    except Exception as exc:
        raise RuntimeError(f"Failed to convert MSG to PDF via Outlook/Word: {source_path.name}") from exc
    finally:
        message = None
        if outlook is not None:
            outlook.Quit()
        pythoncom.CoUninitialize()

    _word_to_pdf(temp_mht, output_path)


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
