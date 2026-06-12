# AGENTS.md — Print Assist

`CLAUDE.md` in this directory is the single source of truth for all coding
agents working on this repo (Codex and Claude Code alike): project context,
architecture map, core invariants, environment constraints, and the
prompt-handoff conventions. Read it before making changes.

The short version:

- Windows-first tkinter app that combines PDFs, images, and Office/Outlook
  files into one A4-fitted, print-ready PDF.
- Preview and final output must come from the same `build_combined_pdf`
  pipeline — never let them diverge.
- Office/MSG conversion is Windows + Microsoft Office only; keep failures
  graceful and off the Tk main thread.
- After behavior changes update `README.md`; after file add/remove/rename
  update `CHATGPT_CONTEXT.md` and the architecture map in `CLAUDE.md`.
