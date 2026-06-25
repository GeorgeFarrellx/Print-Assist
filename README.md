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
- Shows selected files in a tree with vertical and horizontal scrollbars for long file lists/paths.
- Groups each Outlook email's extracted attachments beneath that email; email groups can be collapsed or expanded.
- Sorts the upload list by manual order, filename A–Z, or Outlook email date/time while preserving attachment groups.
- Main controls are grouped into Import, File list/order, and Output/action rows for a clearer layout.
- Reorder file order (Move Up / Move Down).
- Remove selected files or clear all.
- Choose output location/name.
- Create an in-app preview of one combined PDF with A4-fitted pages.
- Save the final combined PDF from the preview window.
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
- Preview is shown in-app with page navigation and zoom, and displays both the overall PDF page number and the current source file's page number and page count.
- Image pages can be cropped interactively in the preview by dragging around the area to keep.
- Outlook email pages can be trimmed at a selected horizontal line; later pages belonging to that same `.msg` are removed while separately listed attachments remain.
- Preview edits are non-destructive and include Undo Edit and Reset Edits controls.
- The current preview page can be deleted without changing or deleting its source file, document, email, or attachment.
- Preview includes a File Summary window listing source file order, type, page range, page count, and full source path.
- "Save Final PDF" copies the exact preview PDF to the selected output path so reviewed pages match saved output.
- Preview should match final output as closely as possible because both use the same combined-PDF generation process.
- Output defaults to: `Print Assist - YYYYMMDD_HHMMSS.pdf`
- If all selected files come from one folder, that folder is used as default save location.
- PDF pages are fitted onto A4 (portrait/landscape chosen automatically), centered, no cropping.
- Images are placed one per A4 page (portrait/landscape chosen automatically), centered, no cropping.
- Word/Excel/MSG files are converted to temporary PDFs first, then added to the combined output PDF.
- `.msg` conversion uses an unattended Outlook-to-Word export styled to closely match Outlook's Memo Style, so large email batches can be previewed without a Microsoft Print to PDF save dialog for every message.
- Adding or dropping a `.msg` automatically places the email first, followed by its visible printable attachments in Outlook order.
- Nested attached `.msg` emails are expanded up to five levels. Hidden inline/signature images and unsupported attachments are skipped.
- Excel output may span multiple pages based on workbook print areas and page setup.
- Folder import reads direct child files only (subfolders are ignored in this version), which is useful for quickly loading files from an Outlook macro's saved Attachments folder.
- Add Client Folder is designed for Outlook macro client folders where a saved `.msg` is in the parent folder and saved attachments are in an `Attachments` (or `Attachment` / `Email Attachments`) subfolder.
- On Windows, whole Outlook email messages and individual Outlook attachments can be dragged directly into the file list; Print Assist saves the dropped virtual item to a temporary file for the current session.

## Current limitations

- Drag-and-drop supports Windows Explorer files/folders and Outlook attachments on Windows, with `tkinterdnd2` used as a fallback when native Windows drag/drop is unavailable.
- If drag/drop is unavailable in the local environment, Add Files, Add Folder, and Add Client Folder remain fully supported.
- Word/Excel/MSG conversion requires installed Microsoft Office/Outlook on Windows.
- Preview-window printing depends on Windows/default PDF print handling behavior.
- Printer selection, duplex, tray selection, paper selection, and other advanced print options are not implemented yet.

## Planned future enhancements

- Direct print button
- Save/load file batches
