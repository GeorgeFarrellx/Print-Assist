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
WORD_HEADER_FOOTER_PRIMARY = 1
WORD_LINE_SPACE_EXACTLY = 4
WORD_ALIGN_PARAGRAPH_CENTER = 1
WORD_FIELD_PAGE = 33
WORD_RELATIVE_POSITION_PAGE = 1
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_FALSE = 0

OUTLOOK_MEMO_FIELD_LABELS = {
    "from:",
    "sent:",
    "to:",
    "cc:",
    "bcc:",
    "subject:",
    "attachments:",
}
OUTLOOK_MEMO_MARGIN = 36.84
OUTLOOK_MEMO_FIELD_TAB = 152.88


def _require_windows() -> None:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception as exc:
        raise RuntimeError("pywin32/COM automation is unavailable on this system") from exc


def _memo_heading_from_message(message: object) -> str:
    recipients = str(getattr(message, "To", "") or "").strip()
    if not recipients:
        return "Outlook Email"
    return recipients.split(";", 1)[0].strip() or "Outlook Email"


def _apply_outlook_memo_style(document: object, heading: str) -> None:
    """Approximate Outlook's Memo Style without invoking a print driver."""
    section = document.Sections(1)
    page_setup = section.PageSetup
    page_setup.LeftMargin = OUTLOOK_MEMO_MARGIN
    page_setup.RightMargin = OUTLOOK_MEMO_MARGIN
    page_setup.TopMargin = 77
    page_setup.BottomMargin = 36
    page_setup.HeaderDistance = 36
    page_setup.FooterDistance = 37.5

    last_field_paragraph = None
    paragraph_limit = min(document.Paragraphs.Count, 12)
    for index in range(1, paragraph_limit + 1):
        paragraph = document.Paragraphs(index)
        text = str(paragraph.Range.Text or "").replace("\r", "").strip()
        label = text.split("\t", 1)[0].strip().lower()
        if label not in OUTLOOK_MEMO_FIELD_LABELS:
            if last_field_paragraph is not None:
                break
            continue

        paragraph.Format.LeftIndent = OUTLOOK_MEMO_FIELD_TAB
        paragraph.Format.FirstLineIndent = -OUTLOOK_MEMO_FIELD_TAB
        paragraph.Format.TabStops.ClearAll()
        paragraph.Format.TabStops.Add(OUTLOOK_MEMO_FIELD_TAB)
        paragraph.Format.LineSpacingRule = WORD_LINE_SPACE_EXACTLY
        paragraph.Format.LineSpacing = 13.2
        paragraph.Range.Font.Name = "Arial"
        paragraph.Range.Font.Size = 10
        last_field_paragraph = paragraph

    if last_field_paragraph is not None:
        last_field_paragraph.Format.SpaceAfter = 14

    header = section.Headers(WORD_HEADER_FOOTER_PRIMARY)
    header.Range.Text = ""
    title = header.Shapes.AddTextbox(
        MSO_TEXT_ORIENTATION_HORIZONTAL,
        OUTLOOK_MEMO_MARGIN,
        48,
        300,
        18,
        header.Range,
    )
    title.RelativeHorizontalPosition = WORD_RELATIVE_POSITION_PAGE
    title.RelativeVerticalPosition = WORD_RELATIVE_POSITION_PAGE
    title.Line.Visible = MSO_FALSE
    title.Fill.Visible = MSO_FALSE
    title.TextFrame.MarginLeft = 0
    title.TextFrame.MarginRight = 0
    title.TextFrame.MarginTop = 0
    title.TextFrame.MarginBottom = 0
    title.TextFrame.TextRange.Text = heading
    title.TextFrame.TextRange.Font.Name = "Arial"
    title.TextFrame.TextRange.Font.Size = 11
    title.TextFrame.TextRange.Font.Bold = True

    rule = header.Shapes.AddLine(
        OUTLOOK_MEMO_MARGIN,
        67.56,
        612 - OUTLOOK_MEMO_MARGIN,
        67.56,
        header.Range,
    )
    rule.RelativeHorizontalPosition = WORD_RELATIVE_POSITION_PAGE
    rule.RelativeVerticalPosition = WORD_RELATIVE_POSITION_PAGE
    rule.Line.Weight = 3.24
    rule.Line.ForeColor.RGB = 0

    footer = section.Footers(WORD_HEADER_FOOTER_PRIMARY)
    footer.Range.Text = ""
    footer.Range.Fields.Add(footer.Range, WORD_FIELD_PAGE)
    footer.Range.Font.Name = "Times New Roman"
    footer.Range.Font.Size = 8
    footer.Range.ParagraphFormat.Alignment = WORD_ALIGN_PARAGRAPH_CENTER


class OfficeConversionSession:
    """Reuse hidden Office applications while building one combined PDF."""

    def __init__(self, temp_dir: Path) -> None:
        self.temp_dir = temp_dir
        self._pythoncom = None
        self._win32_client = None
        self._com_initialised = False
        self._word = None
        self._excel = None
        self._outlook = None
        self._outlook_namespace = None
        self._conversion_index = 0

    def __enter__(self) -> OfficeConversionSession:
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> None:
        _ = exc_type, exc_value, traceback
        self.close()

    def _ensure_com(self) -> None:
        if self._com_initialised:
            return
        _require_windows()
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        self._win32_client = win32com.client
        self._com_initialised = True

    def _get_word(self) -> object:
        self._ensure_com()
        if self._word is None:
            self._word = self._win32_client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = WORD_ALERTS_NONE
        return self._word

    def _get_excel(self) -> object:
        self._ensure_com()
        if self._excel is None:
            self._excel = self._win32_client.DispatchEx("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
        return self._excel

    def _get_outlook_namespace(self) -> object:
        self._ensure_com()
        if self._outlook_namespace is None:
            # Outlook is a single-instance application. Connect to it without
            # quitting the user's open Outlook session when conversion finishes.
            self._outlook = self._win32_client.Dispatch("Outlook.Application")
            self._outlook_namespace = self._outlook.GetNamespace("MAPI")
        return self._outlook_namespace

    def _word_to_pdf(
        self,
        source_path: Path,
        output_path: Path,
        memo_heading: str | None = None,
    ) -> None:
        word = self._get_word()
        doc = None
        try:
            doc = word.Documents.Open(
                str(source_path),
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
            if memo_heading is not None:
                _apply_outlook_memo_style(doc, memo_heading)
            doc.ExportAsFixedFormat(str(output_path), PDF_FORMAT)
        finally:
            if doc is not None:
                doc.Close(False)

    def _excel_to_pdf(self, source_path: Path, output_path: Path) -> None:
        excel = self._get_excel()
        workbook = None
        try:
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

    def _msg_to_pdf(self, source_path: Path, output_path: Path) -> None:
        namespace = self._get_outlook_namespace()
        message = None
        temp_mht = output_path.with_suffix(".mht")
        try:
            message = namespace.OpenSharedItem(str(source_path))
            memo_heading = _memo_heading_from_message(message)
            message.SaveAs(str(temp_mht), OUTLOOK_OL_MHTML)
            self._word_to_pdf(temp_mht, output_path, memo_heading=memo_heading)
        except Exception as exc:
            raise RuntimeError(
                f"Failed to convert Outlook message without printing: {source_path.name}"
            ) from exc
        finally:
            message = None

    def convert(self, source_path: Path, output_path: Path | None = None) -> Path:
        suffix = source_path.suffix.lower()
        if suffix == ".pdf":
            return source_path

        self._conversion_index += 1
        if output_path is None:
            output_path = (
                self.temp_dir
                / f"{self._conversion_index:04d}_{source_path.stem}.converted.pdf"
            )
        output_path.parent.mkdir(parents=True, exist_ok=True)

        if suffix in WORD_EXTENSIONS:
            self._word_to_pdf(source_path, output_path)
            return output_path
        if suffix in EXCEL_EXTENSIONS:
            self._excel_to_pdf(source_path, output_path)
            return output_path
        if suffix in MSG_EXTENSIONS:
            self._msg_to_pdf(source_path, output_path)
            return output_path

        raise ValueError(f"Unsupported Office/MSG file extension: {source_path.suffix}")

    def close(self) -> None:
        self._outlook_namespace = None
        self._outlook = None
        if self._excel is not None:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
        if self._word is not None:
            try:
                self._word.Quit()
            except Exception:
                pass
            self._word = None
        if self._com_initialised:
            self._pythoncom.CoUninitialize()
            self._com_initialised = False


def _word_to_pdf(source_path: Path, output_path: Path) -> None:
    with OfficeConversionSession(output_path.parent) as converter:
        converter._word_to_pdf(source_path, output_path)


def _excel_to_pdf(source_path: Path, output_path: Path) -> None:
    with OfficeConversionSession(output_path.parent) as converter:
        converter._excel_to_pdf(source_path, output_path)


def _msg_to_pdf(source_path: Path, output_path: Path, temp_dir: Path) -> None:
    with OfficeConversionSession(temp_dir) as converter:
        converter._msg_to_pdf(source_path, output_path)


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
