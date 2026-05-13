# Print Assist

Print Assist is a Windows-first Python desktop app that combines mixed printable files into one clean, print-ready PDF.

## Supported file types (v1)

- PDF
- JPG / JPEG
- PNG
- BMP
- TIF / TIFF

## What it does

- Add multiple files using file picker.
- Shows selected files in a list.
- Reorder file order (Move Up / Move Down).
- Remove selected files or clear all.
- Choose output location/name.
- Create one combined PDF with A4-fitted pages.
- Open the generated PDF after creation.
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

## Output behavior

- Output defaults to: `Print Assist - YYYYMMDD_HHMMSS.pdf`
- If all selected files come from one folder, that folder is used as default save location.
- PDF pages are fitted onto A4 (portrait/landscape chosen automatically), centered, no cropping.
- Images are placed one per A4 page (portrait/landscape chosen automatically), centered, no cropping.

## Current limitations

- Drag-and-drop support is best-effort and may depend on local Tk setup.
- Word, Excel, and Outlook `.msg` are not included in v1.
- No direct printing yet (output is combined PDF only).

## Planned future enhancements

- Word support
- Excel support
- Outlook `.msg` support
- Direct print button
- Save/load file batches
- Integration with Outlook attachment saving workflow
