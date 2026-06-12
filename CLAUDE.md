# Claude Code Context — Print Assist

This is the working context file for coding agents (Claude Code / Codex) on this
repository. It is the agent-side counterpart to `CHATGPT_CONTEXT.md` (which is a
link list given to chat models such as ChatGPT or Claude when refining prompts).

Use this file two ways:

1. **When implementing a refined prompt** that came from ChatGPT or Claude chat:
   treat the prompt as intent, and this file plus the actual code as truth.
2. **When writing a refined prompt** for Claude Code or Codex: include the
   relevant facts below so the implementing agent does not have to rediscover them.

## What the program is

Print Assist is a Windows-first Python desktop app (tkinter) that combines mixed
printable files — PDFs, images, Word/Excel documents, and Outlook `.msg` emails —
into one clean, print-ready PDF with every page fitted onto A4.

- Entry point: `main.py` (uses a `tkinterdnd2` drag-and-drop root when available,
  falls back to plain `tk.Tk`), which calls `run()` in `print_assist/app.py`.
- Dependencies: PyMuPDF (fitz), Pillow, tkinterdnd2, pywin32 (`requirements.txt`).
- Distribution: GitHub Actions workflow `.github/workflows/build-exe.yml` builds a
  single-file Windows EXE with PyInstaller (manual `workflow_dispatch` trigger).

## Architecture map

| File | Responsibility |
| --- | --- |
| `main.py` | Creates the Tk root (drag-and-drop capable if possible) and starts the app. |
| `print_assist/app.py` | `PrintAssistApp` — main window. Import buttons (Add Files / Add Folder / Add Client Folder), drag-and-drop, file list with reorder/remove/clear, output path selection, and preview generation on a background worker thread polled via a queue and `root.after`. |
| `print_assist/file_utils.py` | Supported-extension filtering, folder scanning (non-recursive), client-folder scanning (parent files plus `Attachments` / `Attachment` / `Email Attachments` child folders), default output path logic. |
| `print_assist/office_converter.py` | Windows-only COM automation via pywin32: Word → PDF, Excel → PDF, and `.msg` → MHT (Outlook) → PDF (Word). `convert_to_pdf_if_needed` is the single entry point. |
| `print_assist/pdf_builder.py` | `build_combined_pdf` — combines everything into one PDF with PyMuPDF. A4 portrait/landscape chosen per page, content centered, never cropped. Produces a source manifest (which source file produced which page range). |
| `print_assist/preview_window.py` | `PreviewWindow` — in-app preview with page navigation, zoom, per-page source file name, File Summary window, Save Final PDF, open externally, and print. |

## Core invariants — any change must respect these

- **Preview equals output.** The preview PDF and the final PDF come from the same
  `build_combined_pdf` pipeline. "Save Final PDF" copies the exact preview file;
  "Print Preview PDF" prints the exact preview file. Never introduce a second
  rendering path that could make saved/printed output differ from the preview.
- **A4 fitting.** Every page/image is fitted onto A4 (orientation chosen
  automatically), centered, with no cropping.
- **UI stays responsive.** Long work (conversion, combining) runs on a background
  thread; results flow back through a queue polled with `root.after`. Do not call
  Office/PDF conversion on the Tk main thread.
- **Graceful degradation.** Office/Outlook conversion and drag-and-drop are
  optional capabilities: missing Office, non-Windows platforms, or a missing
  `tkinterdnd2` must produce clear errors or fallbacks, not crashes.
- **Folder imports are non-recursive** (direct children only), by design — they
  match an Outlook macro's saved-attachments folder layout.

## Environment constraints for agents

- Word/Excel/MSG conversion requires installed Microsoft Office/Outlook on
  Windows; it cannot run in a Linux/cloud agent environment. Neither can the
  tkinter UI without a display. So in a cloud session, verify changes by review,
  import checks, and any logic that can run headless — and say plainly what was
  and was not executed.
- Keep changes compatible with the PyInstaller one-file build (avoid data files
  loaded by relative path, dynamic imports PyInstaller cannot see, etc.).

## Writing refined prompts for Claude Code / Codex

When turning an idea into an implementation prompt, include:

1. **Goal in user terms** — what the user sees/does differently afterwards.
2. **Exact files involved** — use the architecture map above; name modules, not
   "the converter code".
3. **Invariants that apply** — copy the relevant bullets from "Core invariants"
   into the prompt so they survive the handoff.
4. **What done looks like** — observable behavior, plus what to update in
   `README.md` (it documents behavior in detail and is kept current).
5. **What NOT to change** — e.g. "do not touch the preview/save pipeline".

Keep prompts self-contained: the implementing agent may not have this repo's chat
history, only the repo itself.

## Implementing prompts that came from ChatGPT / Claude chat

- Chat models read this repo through the raw links in `CHATGPT_CONTEXT.md`, which
  point at `main` — their view may be **stale or wrong** about line numbers,
  function names, or current behavior. Verify every claim against the actual code
  before editing; follow the prompt's intent, not its possibly outdated specifics.
- If the prompt conflicts with a core invariant above, stop and flag it instead
  of implementing it.
- Work on a feature branch, commit with clear messages, push, and open a PR.

## Housekeeping after changes

- Adding, removing, or renaming source files → update `CHATGPT_CONTEXT.md`
  (the raw-link list) and the architecture map in this file.
- Behavior changes → update `README.md`, which is the detailed behavior spec.
- New third-party imports → update `requirements.txt`.
