from __future__ import annotations

import os
import queue
import subprocess
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES
except Exception:
    DND_FILES = None

from . import APP_NAME
from .file_utils import SUPPORTED_EXTENSIONS, default_output_path, filter_supported_files, get_supported_files_from_client_folder, get_supported_files_from_folder
from .pdf_builder import build_combined_pdf
from .preview_window import PreviewWindow


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
        self.file_count_var = tk.StringVar(value="Selected files: 0")
        self._preview_running = False
        self._preview_queue: queue.Queue[tuple[str, object]] | None = None
        self._preview_temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        title = ttk.Label(self.root, text=APP_NAME, font=("Segoe UI", 18, "bold"))
        title.pack(pady=(12, 8))

        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        list_label = ttk.Label(frame, text="Drop files/folders here, or use Add Files / Add Folder:")
        list_label.pack(anchor="w")

        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 10))

        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=18, xscrollcommand=None, yscrollcommand=None)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        y_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.listbox.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.listbox.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        drag_drop_enabled = False
        if DND_FILES is not None and hasattr(self.root, "drop_target_register") and hasattr(self.root, "dnd_bind"):
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]
                drag_drop_enabled = True
            except Exception:
                drag_drop_enabled = False
        if not drag_drop_enabled:
            self.status_var.set("Use Add Files or Add Folder")

        controls = ttk.Frame(frame)
        controls.pack(fill=tk.X, pady=(0, 8))

        button_groups = [
            [("Add Files", self.add_files), ("Add Folder", self.add_folder), ("Add Client Folder", self.add_client_folder)],
            [("Remove Selected", self.remove_selected), ("Move Up", self.move_up), ("Move Down", self.move_down), ("Clear", self.clear_files)],
            [("Choose Output", self.choose_output), ("Preview Print Assist PDF", self.create_preview), ("Open Output Folder", self.open_output_folder)],
        ]

        self.buttons: dict[str, ttk.Button] = {}
        for row_idx, group in enumerate(button_groups):
            row_frame = ttk.Frame(controls)
            row_frame.pack(fill=tk.X, pady=2)
            for col_idx, (label, command) in enumerate(group):
                button = ttk.Button(row_frame, text=label, command=command)
                button.grid(row=0, column=col_idx, padx=4, pady=2, sticky="w")
                self.buttons[label] = button

        ttk.Label(frame, textvariable=self.file_count_var).pack(anchor="w", pady=(0, 2))
        ttk.Label(frame, textvariable=self.output_var).pack(anchor="w", pady=(6, 2))
        ttk.Progressbar(frame, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=2)
        ttk.Label(frame, textvariable=self.status_var).pack(anchor="w", pady=(6, 0))

    def _on_drop(self, event: tk.Event) -> None:
        raw = self.root.tk.splitlist(event.data)
        dropped_files: list[str] = []
        unsupported: list[Path] = []

        for item in raw:
            path = Path(item)
            if path.is_dir():
                supported, folder_unsupported = get_supported_files_from_folder(path)
                dropped_files.extend(str(p) for p in supported)
                unsupported.extend(folder_unsupported)
            else:
                dropped_files.append(item)

        self._append_paths(dropped_files, precomputed_unsupported=unsupported)

    def add_files(self) -> None:
        types = [("PDF", "*.pdf"), ("Images", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"), ("Word documents", "*.doc *.docx"), ("Excel workbooks", "*.xls *.xlsx *.xlsm *.xlsb"), ("Outlook messages", "*.msg"), ("All files", "*.*")]
        selected = filedialog.askopenfilenames(title="Select files", filetypes=types)
        self._append_paths(selected)

    def add_folder(self) -> None:
        selected_folder = filedialog.askdirectory(title="Select folder")
        if not selected_folder:
            return

        supported, unsupported = get_supported_files_from_folder(Path(selected_folder))
        added = 0
        for p in supported:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, str(p))
                added += 1

        if added and self.output_path is None:
            self.output_path = default_output_path(self.files)
            self.output_var.set(f"Output: {self.output_path}")

        if added:
            self._update_file_count()

        if unsupported:
            display_unsupported = unsupported[:20]
            warning_msg = "Unsupported files skipped:\n" + "\n".join(u.name for u in display_unsupported)
            remaining = len(unsupported) - len(display_unsupported)
            if remaining > 0:
                warning_msg += f"\n...and {remaining} more unsupported file(s)."
            messagebox.showwarning(APP_NAME, warning_msg)

        if added:
            self.status_var.set(f"Added {added} file(s) from folder.")
        elif supported:
            self.status_var.set("No new files were added from folder.")
        else:
            self.status_var.set("No supported files found in the selected folder.")


    def add_client_folder(self) -> None:
        selected_folder = filedialog.askdirectory(title="Select client folder")
        if not selected_folder:
            return

        supported, unsupported = get_supported_files_from_client_folder(Path(selected_folder))
        added = 0
        for p in supported:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, str(p))
                added += 1

        if added and self.output_path is None:
            self.output_path = default_output_path(self.files)
            self.output_var.set(f"Output: {self.output_path}")

        if added:
            self._update_file_count()

        if unsupported:
            display_unsupported = unsupported[:20]
            warning_msg = "Unsupported files skipped:\n" + "\n".join(u.name for u in display_unsupported)
            remaining = len(unsupported) - len(display_unsupported)
            if remaining > 0:
                warning_msg += f"\n...and {remaining} more unsupported file(s)."
            messagebox.showwarning(APP_NAME, warning_msg)

        if added:
            self.status_var.set(f"Added {added} file(s) from client folder.")
        elif supported:
            self.status_var.set("No new files were added from client folder.")
        else:
            self.status_var.set("No supported files found in the selected client folder.")

    def _append_paths(self, raw_paths: tuple[str, ...] | list[str], precomputed_unsupported: list[Path] | None = None) -> None:
        supported, unsupported = filter_supported_files(raw_paths)
        if precomputed_unsupported:
            unsupported = list(unsupported) + precomputed_unsupported
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
            display_unsupported = unsupported[:20]
            warning_msg = "Unsupported files skipped:\n" + "\n".join(u.name for u in display_unsupported)
            remaining = len(unsupported) - len(display_unsupported)
            if remaining > 0:
                warning_msg += f"\n...and {remaining} more unsupported file(s)."
            messagebox.showwarning(APP_NAME, warning_msg)

        self._update_file_count()
        self.status_var.set(f"{len(self.files)} file(s) selected.")

    def _update_file_count(self) -> None:
        self.file_count_var.set(f"Selected files: {len(self.files)}")

    def remove_selected(self) -> None:
        indices = list(self.listbox.curselection())
        for i in reversed(indices):
            self.listbox.delete(i)
            self.files.pop(i)
        self._update_file_count()
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
        self._update_file_count()
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

    def _set_preview_controls_enabled(self, enabled: bool) -> None:
        labels = [
            "Add Files",
            "Add Folder",
            "Add Client Folder",
            "Remove Selected",
            "Move Up",
            "Move Down",
            "Clear",
            "Choose Output",
            "Preview Print Assist PDF",
        ]
        state = tk.NORMAL if enabled else tk.DISABLED
        for label in labels:
            button = self.buttons.get(label)
            if button is not None:
                button.configure(state=state)

    def _cleanup_preview_temp_dir(self) -> None:
        if self._preview_temp_dir_obj is not None:
            self._preview_temp_dir_obj.cleanup()
            self._preview_temp_dir_obj = None

    def _poll_preview_queue(self) -> None:
        if self._preview_queue is None:
            return

        while True:
            try:
                event_type, payload = self._preview_queue.get_nowait()
            except queue.Empty:
                break

            if event_type == "progress":
                current, total, file_path = payload
                percent = 0 if total == 0 else (current / total) * 100
                self.progress_var.set(percent)
                self.status_var.set(f"Processing {current} of {total}: {Path(file_path).name}")
            elif event_type == "done":
                processed, warnings, preview_pdf = payload
                self._preview_running = False
                self._set_preview_controls_enabled(True)
                self.progress_var.set(100)
                if warnings:
                    messagebox.showwarning(APP_NAME, "\n".join(warnings))
                if not processed:
                    self._cleanup_preview_temp_dir()
                    self.progress_var.set(0)
                    self.status_var.set("Error")
                    return
                self.status_var.set("Preview ready")
                preview_window = PreviewWindow(
                    parent=self.root,
                    preview_pdf_path=preview_pdf,
                    output_path=self.output_path,
                    open_pdf_callback=self.open_pdf,
                    on_status_change=self.status_var.set,
                    on_close_callback=self._cleanup_preview_temp_dir,
                )
                _ = preview_window
                return
            elif event_type == "error":
                self._preview_running = False
                self._set_preview_controls_enabled(True)
                self._cleanup_preview_temp_dir()
                self.progress_var.set(0)
                self.status_var.set("Error")
                messagebox.showerror(APP_NAME, payload)
                return

        if self._preview_running:
            self.root.after(100, self._poll_preview_queue)

    def create_preview(self) -> None:
        if self._preview_running:
            return

        if not self.files:
            messagebox.showerror(APP_NAME, "Please add at least one supported file.")
            return

        if self.output_path is None:
            self.output_path = default_output_path(self.files)
            self.output_var.set(f"Output: {self.output_path}")

        self._cleanup_preview_temp_dir()
        self._preview_temp_dir_obj = tempfile.TemporaryDirectory(prefix="print_assist_preview_")
        preview_dir = Path(self._preview_temp_dir_obj.name)
        preview_pdf = preview_dir / "preview.pdf"

        self._preview_running = True
        self._set_preview_controls_enabled(False)
        self.status_var.set("Creating preview...")
        self.progress_var.set(0)
        self._preview_queue = queue.Queue()

        files_to_process = list(self.files)

        def progress_callback(current: int, total: int, file_path: Path, status_text: str) -> None:
            _ = status_text
            if self._preview_queue is not None:
                self._preview_queue.put(("progress", (current, total, str(file_path))))

        def worker() -> None:
            try:
                processed, warnings = build_combined_pdf(files_to_process, preview_pdf, progress_callback=progress_callback)
                if self._preview_queue is not None:
                    self._preview_queue.put(("done", (processed, warnings, preview_pdf)))
            except Exception as exc:
                if self._preview_queue is not None:
                    self._preview_queue.put(("error", f"Failed to create preview PDF:\n{exc}"))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self._poll_preview_queue)

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


def run(root: tk.Tk | None = None) -> None:
    root = root if root is not None else tk.Tk()
    app = PrintAssistApp(root)
    _ = app
    root.mainloop()
