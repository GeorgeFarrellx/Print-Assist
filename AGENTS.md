# Agent Instructions — Print Assist

Read `CLAUDE.md` in this directory: it is the single source of truth for project
context, architecture, core invariants, and prompt-handling conventions for all
coding agents (Codex and Claude Code alike). Follow it.

Quick essentials:

- Windows-first tkinter app that combines PDFs/images/Office files into one
  A4-fitted, print-ready PDF.
- Preview and final output must come from the same `build_combined_pdf` pipeline.
- Office/Outlook conversion is Windows + Microsoft Office only — keep failures
  graceful and never run conversions on the Tk main thread.
- After behavior changes update `README.md`; after file add/remove/rename update
  `CHATGPT_CONTEXT.md` and the architecture map in `CLAUDE.md`.
