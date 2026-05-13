from __future__ import annotations

import os
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import APP_NAME
from .file_utils import SUPPORTED_EXTENSIONS, default_output_path, filter_supported_files
from .pdf_builder import build_combined_pdf


class PrintAssistApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("900x560")

        self.files: list[Path] = []
        self.output_path: Path | None = None

        self.status_var = tk.StringVar(value="Ready")
        self.output_var = tk.StringVar(value="No output file selected")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()

    def _build_ui(self) -> None:
        title = ttk.Label(self.root, text=APP_NAME, font=("Segoe UI", 18, "bold"))
        title.pack(pady=(12, 8))

        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        list_label = ttk.Label(frame, text="Drop files here (or use Add Files):")
        list_label.pack(anchor="w")

        self.listbox = tk.Listbox(frame, selectmode=tk.EXTENDED, height=18)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=(6, 10))

        try:
            self.root.drop_target_register(tk.DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]
        except Exception:
            pass

        controls = ttk.Frame(frame)
        controls.pack(fill=tk.X, pady=(0, 8))

        buttons = [
            ("Add Files", self.add_files),
            ("Remove Selected", self.remove_selected),
            ("Move Up", self.move_up),
            ("Move Down", self.move_down),
            ("Clear", self.clear_files),
            ("Choose Output", self.choose_output),
            ("Create Print Assist PDF", self.create_pdf),
            ("Open Output Folder", self.open_output_folder),
        ]

        for idx, (label, command) in enumerate(buttons):
            ttk.Button(controls, text=label, command=command).grid(row=0, column=idx, padx=4, pady=2)

        ttk.Label(frame, textvariable=self.output_var).pack(anchor="w", pady=(6, 2))
        ttk.Progressbar(frame, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=2)
        ttk.Label(frame, textvariable=self.status_var).pack(anchor="w", pady=(6, 0))

    def _on_drop(self, event: tk.Event) -> None:
        raw = self.root.tk.splitlist(event.data)
        self._append_paths(raw)

    def add_files(self) -> None:
        types = [("PDF", "*.pdf"), ("Images", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"), ("Word documents", "*.doc *.docx"), ("Excel workbooks", "*.xls *.xlsx *.xlsm *.xlsb"), ("Outlook messages", "*.msg"), ("All files", "*.*")]
        selected = filedialog.askopenfilenames(title="Select files", filetypes=types)
        self._append_paths(selected)

    def _append_paths(self, raw_paths: tuple[str, ...] | list[str]) -> None:
        supported, unsupported = filter_supported_files(raw_paths)
        added = 0
        for p in supported:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, str(p))
                added += 1

        if added and self.output_path is None:
            self.output_path = default_output_path(self.files)
            self.output_var.set(f"Output: {self.output_path}")

        if unsupported:
            warning_msg = "Unsupported files skipped:\n" + "\n".join(u.name for u in unsupported)
            messagebox.showwarning(APP_NAME, warning_msg)

        self.status_var.set(f"{len(self.files)} file(s) selected.")

    def remove_selected(self) -> None:
        indices = list(self.listbox.curselection())
        for i in reversed(indices):
            self.listbox.delete(i)
            self.files.pop(i)
        self.status_var.set(f"{len(indices)} file(s) removed.")

    def move_up(self) -> None:
        for i in self.listbox.curselection():
            if i == 0:
                continue
            self.files[i - 1], self.files[i] = self.files[i], self.files[i - 1]
            self._refresh_listbox(select_index=i - 1)

    def move_down(self) -> None:
        for i in reversed(self.listbox.curselection()):
            if i >= len(self.files) - 1:
                continue
            self.files[i + 1], self.files[i] = self.files[i], self.files[i + 1]
            self._refresh_listbox(select_index=i + 1)

    def _refresh_listbox(self, select_index: int | None = None) -> None:
        self.listbox.delete(0, tk.END)
        for p in self.files:
            self.listbox.insert(tk.END, str(p))
        if select_index is not None:
            self.listbox.select_set(select_index)

    def clear_files(self) -> None:
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self.progress_var.set(0)
        self.status_var.set("File list cleared.")

    def choose_output(self) -> None:
        initial = str(self.output_path) if self.output_path else str(default_output_path(self.files))
        selected = filedialog.asksaveasfilename(
            title="Choose output PDF",
            defaultextension=".pdf",
            initialfile=Path(initial).name,
            initialdir=str(Path(initial).parent),
            filetypes=[("PDF files", "*.pdf")],
        )
        if selected:
            self.output_path = Path(selected)
            self.output_var.set(f"Output: {self.output_path}")
            self.status_var.set("Output path selected.")

    def create_pdf(self) -> None:
        if not self.files:
            messagebox.showerror(APP_NAME, "Please add at least one supported file.")
            return

        if self.output_path is None:
            self.output_path = default_output_path(self.files)

        try:
            self.status_var.set("Creating combined PDF...")
            self.progress_var.set(30)
            processed, warnings = build_combined_pdf(self.files, self.output_path)
            self.progress_var.set(100)

            if warnings:
                messagebox.showwarning(APP_NAME, "\n".join(warnings))

            messagebox.showinfo(APP_NAME, f"Created PDF with {len(processed)} file(s).\n{self.output_path}")
            self.status_var.set("Done")
            self.open_pdf(self.output_path)
        except Exception as exc:
            self.progress_var.set(0)
            self.status_var.set("Error")
            messagebox.showerror(APP_NAME, f"Failed to create combined PDF:\n{exc}")

    def open_output_folder(self) -> None:
        if self.output_path and self.output_path.parent.exists():
            os.startfile(self.output_path.parent)  # type: ignore[attr-defined]
        else:
            messagebox.showinfo(APP_NAME, "No output folder available yet.")

    def open_pdf(self, path: Path) -> None:
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception:
            subprocess.run(["xdg-open", str(path)], check=False)


def run() -> None:
    root = tk.Tk()
    app = PrintAssistApp(root)
    _ = app
    root.mainloop()
