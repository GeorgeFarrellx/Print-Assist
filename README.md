# Print Assist

Print Assist is a Windows-first Python desktop app that combines mixed printable files into one clean, print-ready PDF.

## Supported file types (current)

- PDF
- JPG / JPEG
- PNG
- BMP
- TIF / TIFF
- DOC / DOCX (Microsoft Word required on Windows)
- XLS / XLSX / XLSM / XLSB (Microsoft Excel required on Windows)
- MSG (Microsoft Outlook required on Windows)

## What it does

- Add multiple files using file picker.
- Shows selected files in a list.
- Reorder file order (Move Up / Move Down).
- Remove selected files or clear all.
- Choose output location/name.
- Create an in-app preview of one combined PDF with A4-fitted pages.
- Save the final combined PDF from the preview window.
- Print the preview PDF from the preview window.
- Open preview or saved PDF externally.
- Open output folder.

## Install

1. Create/activate a virtual environment (recommended).
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

## Preview and output behavior

- "Preview Print Assist PDF" generates a temporary combined preview PDF using the same build pipeline as final output.
- Preview is shown in-app with page navigation and zoom.
- "Save Final PDF" copies the exact preview PDF to the selected output path so reviewed pages match saved output.
- "Print Preview PDF" sends the exact generated preview PDF to Windows/default PDF print handling after confirmation.
- Preview should match final output as closely as possible because both use the same combined-PDF generation process.
- Output defaults to: `Print Assist - YYYYMMDD_HHMMSS.pdf`
- If all selected files come from one folder, that folder is used as default save location.
- PDF pages are fitted onto A4 (portrait/landscape chosen automatically), centered, no cropping.
- Images are placed one per A4 page (portrait/landscape chosen automatically), centered, no cropping.
- Word/Excel/MSG files are converted to temporary PDFs first, then added to the combined output PDF.
- `.msg` conversion uses Outlook to export email header/body content to a temporary MHT/MHTML file, then uses Word to convert that file to PDF.
- Excel output may span multiple pages based on workbook print areas and page setup.
- `.msg` conversion includes email message content only; embedded `.msg` attachments are not automatically included and should be added separately as saved files.

## Current limitations

- Drag-and-drop support is best-effort and may depend on local Tk setup.
- Word/Excel/MSG conversion requires installed Microsoft Office/Outlook on Windows.
- Preview-window printing depends on Windows/default PDF print handling behavior.
- Printer selection, duplex, tray selection, paper selection, and other advanced print options are not implemented yet.

## Planned future enhancements

- Direct print button
- Save/load file batches
- Integration with Outlook attachment saving workflow
