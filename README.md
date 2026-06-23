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
- Add a folder and import supported files from that folder (non-recursive).
- Add Client Folder to import supported direct files plus supported files from direct child Attachments/Attachment/Email Attachments folders (non-recursive beyond those folders).
- Shows selected files in a list with vertical and horizontal scrollbars for long file lists/paths.
- Main controls are grouped into Import, File list/order, and Output/action rows for a clearer layout.
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
- Preview generation runs in a background worker so the main window stays responsive during longer conversions.
- Preview progress now updates per file with status text showing which file is being processed.
- Preview is shown in-app with page navigation and zoom, and displays the source file name for the current page when available.
- Preview includes a File Summary window listing source file order, type, page range, page count, and full source path.
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
- Folder import reads direct child files only (subfolders are ignored in this version), which is useful for quickly loading files from an Outlook macro's saved Attachments folder.
- Add Client Folder is designed for Outlook macro client folders where a saved `.msg` is in the parent folder and saved attachments are in an `Attachments` (or `Attachment` / `Email Attachments`) subfolder.
- On Windows, Outlook attachments can be dragged directly from the Outlook app into the file list; Print Assist saves the dropped virtual attachment to a temporary file for the current session.

## Current limitations

- Drag-and-drop supports Windows Explorer files/folders and Outlook attachments on Windows, with `tkinterdnd2` used as a fallback when native Windows drag/drop is unavailable.
- If drag/drop is unavailable in the local environment, Add Files, Add Folder, and Add Client Folder remain fully supported.
- Word/Excel/MSG conversion requires installed Microsoft Office/Outlook on Windows.
- Preview-window printing depends on Windows/default PDF print handling behavior.
- Printer selection, duplex, tray selection, paper selection, and other advanced print options are not implemented yet.

## Planned future enhancements

- Direct print button
- Save/load file batches
