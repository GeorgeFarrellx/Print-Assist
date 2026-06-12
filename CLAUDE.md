# CLAUDE.md — Print Assist agent context

This file is the coding-agent reference for this repository. It complements
`CHATGPT_CONTEXT.md`, which is a list of raw GitHub links given to chat models
(ChatGPT, Claude on the web) when they help refine prompts.

The workflow this file supports:

1. The user asks a chat model (ChatGPT / Claude) to refine an idea into a prompt.
2. That prompt is handed to a coding agent (Claude Code / Codex) to implement.

So this file is used in two directions:

- **Implementing a refined prompt:** treat the prompt as a statement of intent.
  This file and the actual code are the ground truth. Chat models read the repo
  through the raw links in `CHATGPT_CONTEXT.md`, which point at `main`, so their
  view can be stale — verify file names, function names, and behavior claims
  against the real code before editing, and follow the intent rather than any
  outdated specifics.
- **Writing a refined prompt** for Claude Code or Codex: pull the relevant facts
  from this file into the prompt so the implementing agent gets them up front
  (see "Writing prompts for coding agents" below).

## What Print Assist is

A Windows-first Python desktop app (tkinter) that combines mixed printable
files — PDFs, images (JPG/PNG/BMP/TIFF), Word/Excel documents, and Outlook
`.msg` emails — into a single print-ready PDF with every page fitted onto A4.

- Run with `python main.py`. Dependencies: PyMuPDF (`fitz`), Pillow,
  `tkinterdnd2` (optional drag-and-drop), `pywin32` (Office COM automation).
- A GitHub Actions workflow (`.github/workflows/build-exe.yml`) builds a
  one-file Windows EXE with PyInstaller, triggered manually.
- `README.md` is the detailed behavior spec and is kept current.

## Architecture map

| File | Responsibility |
| --- | --- |
| `main.py` | Creates the Tk root — `TkinterDnD.Tk()` when `tkinterdnd2` is available, plain `tk.Tk()` otherwise — and calls `run()` from `print_assist.app`. |
| `print_assist/__init__.py` | Defines `APP_NAME` ("Print Assist"). |
| `print_assist/app.py` | `PrintAssistApp`, the main window: Add Files / Add Folder / Add Client Folder, drag-and-drop, file list with reorder/remove/clear, output path selection, and preview generation. Preview runs `build_combined_pdf` on a `threading.Thread` worker; progress/done/error events flow through a `queue.Queue` polled with `root.after(100, ...)`. |
| `print_assist/file_utils.py` | `SUPPORTED_EXTENSIONS`, file filtering, non-recursive folder scanning, client-folder scanning (direct files plus `Attachments` / `Attachment` / `Email Attachments` child folders), and `default_output_path` (timestamped name; saved into the common parent folder when all inputs share one). |
| `print_assist/office_converter.py` | Windows-only COM automation via pywin32. `convert_to_pdf_if_needed(source, temp_dir)` is the single entry point: Word → PDF, Excel → PDF, and `.msg` → MHT (Outlook) → PDF (Word). Raises a clear `RuntimeError` when COM/Office is unavailable. |
| `print_assist/pdf_builder.py` | `build_combined_pdf(files, output_path, progress_callback, manifest_callback)`: fits each PDF page or image onto an A4 page (portrait/landscape chosen per page, centered, 24pt margin, never cropped), converts Office/MSG files via the converter first, collects per-file warnings instead of aborting, and emits a manifest entry (source path/name/extension, output page range) per file. |
| `print_assist/preview_window.py` | `PreviewWindow`: page navigation, zoom, per-page source file name, File Summary window, Save Final PDF (copies the preview file), open externally, print. |

## Core invariants — changes must respect these

- **Preview equals output.** Preview and final PDF come from the same
  `build_combined_pdf` pipeline; Save Final PDF copies the exact preview file
  and Print sends that exact file. Never add a second rendering path that could
  make saved or printed output differ from what was previewed.
- **A4 fitting, no cropping.** Every page and image lands on an A4 page,
  orientation chosen automatically, content centered.
- **The Tk main thread stays free.** Conversion and PDF building run on a
  worker thread; results return through the queue + `root.after` polling. Never
  run Office COM conversion or `build_combined_pdf` on the main thread.
- **Graceful degradation.** Missing Office/Outlook, non-Windows platforms, and
  missing `tkinterdnd2` must produce clear error messages or fallbacks, never
  crashes. Per-file failures in `build_combined_pdf` become warnings, not
  aborts.
- **Folder imports are intentionally non-recursive** (direct children only) —
  the layout matches an Outlook macro's saved-attachments folders. Don't "fix"
  this by adding recursion.

## Environment constraints for agents

- Office/Outlook conversion needs Microsoft Office on Windows, and the tkinter
  UI needs a display. Neither runs in a headless Linux/cloud session, so verify
  changes there by code review, import checks, and whatever logic runs headless
  — and state plainly what was and was not actually executed.
- Keep changes PyInstaller-one-file friendly: no data files resolved by
  relative path at runtime, no dynamic imports PyInstaller can't detect.

## Writing prompts for coding agents (Claude Code / Codex)

A refined prompt should be self-contained — the implementing agent has the repo
but not the chat conversation. Include:

1. **The goal in user-visible terms** — what works differently afterwards.
2. **The exact files involved**, named from the architecture map above.
3. **The invariants that apply**, copied from the list above so they survive
   the handoff.
4. **Definition of done** — observable behavior, plus the documentation updates
   (usually `README.md`).
5. **Explicit non-goals** — e.g. "do not touch the preview/save pipeline".

## Implementing prompts from chat models

- Verify the prompt's claims against the current code first; the chat model may
  have read stale `main` via `CHATGPT_CONTEXT.md`.
- If a prompt conflicts with a core invariant, stop and flag the conflict
  instead of implementing it.
- Work on a feature branch, commit with clear messages, push, and open a PR.

## Housekeeping after changes

- Behavior changed → update `README.md` (it documents behavior in detail).
- Files added/removed/renamed → update the raw-link list in
  `CHATGPT_CONTEXT.md` and the architecture map here.
- New third-party dependency → update `requirements.txt` (and check the
  PyInstaller build still covers it).
