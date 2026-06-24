from __future__ import annotations

import os
import queue
import subprocess
import tempfile
import threading
import tkinter as tk
from collections.abc import Iterable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES
except Exception:
    DND_FILES = None

from . import APP_NAME
from .file_utils import (
    PRINTABLE_EXTENSIONS,
    ZIP_EXTENSIONS,
    default_output_path,
    filter_supported_files,
    get_supported_files_from_client_folder,
    get_supported_files_from_folder,
)
from .outlook_message import extract_msg_attachments, safe_outlook_attachment_name, unique_file_path
from .pdf_builder import build_combined_pdf
from .preview_window import PreviewWindow
from .windows_drop import (
    NativeWindowsDropTarget,
    WINDOWS_NATIVE_DROP_AVAILABLE,
    ole_initialize,
    ole_uninitialize,
    register_drop_target,
    revoke_drop_target,
)
from .zip_renamer import ZipExtractionWarning, default_extracted_folder_path, rename_and_extract_zip_contents, unique_folder_path


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
        self._outlook_drop_temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None
        self._native_windows_drop_hwnd: int | None = None
        self._native_windows_drop_target: NativeWindowsDropTarget | None = None
        self._native_windows_drop_target_com: object | None = None
        self._native_drop_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._windows_ole_initialized = ole_initialize()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self.root.after(50, self._poll_native_drop_queue)

    def _build_ui(self) -> None:
        title = ttk.Label(self.root, text=APP_NAME, font=("Segoe UI", 18, "bold"))
        title.pack(pady=(12, 8))

        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        list_label = ttk.Label(frame, text="Drop files, folders, or Outlook email messages here:")
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

        drag_drop_enabled = self._wire_drag_drop()
        if not drag_drop_enabled:
            self.status_var.set("Use Add Files, Add Folder, or Add Client Folder")

        controls = ttk.Frame(frame)
        controls.pack(fill=tk.X, pady=(0, 8))

        button_groups = [
            [("Add Files", self.add_files), ("Add Folder", self.add_folder), ("Add Client Folder", self.add_client_folder)],
            [("Remove Selected", self.remove_selected), ("Move Up", self.move_up), ("Move Down", self.move_down), ("Clear", self.clear_files)],
            [("Choose Output", self.choose_output), ("Preview Print Assist PDF", self.create_preview), ("Rename + Extract ZIP", self.rename_zip_contents), ("Open Output Folder", self.open_output_folder)],
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

    def _wire_drag_drop(self) -> bool:
        if self._wire_native_windows_drop_target():
            return True
        return self._wire_tkinter_drop_target()

    def _wire_native_windows_drop_target(self) -> bool:
        if not WINDOWS_NATIVE_DROP_AVAILABLE or not self._windows_ole_initialized:
            return False
        try:
            self.root.update_idletasks()
            hwnd = int(self.listbox.winfo_id())
            target = NativeWindowsDropTarget(
                on_paths=self._queue_native_drop_paths,
                materialise_virtual_files=self._materialise_outlook_virtual_attachments,
                on_error=self._queue_native_drop_error,
            )
            wrapped = register_drop_target(hwnd, target)
        except Exception:
            self._teardown_native_windows_drop_target()
            return False

        if wrapped is None:
            return False
        self._native_windows_drop_hwnd = hwnd
        self._native_windows_drop_target = target
        self._native_windows_drop_target_com = wrapped
        return True

    def _wire_tkinter_drop_target(self) -> bool:
        if DND_FILES is None or not hasattr(self.root, "drop_target_register") or not hasattr(self.root, "dnd_bind"):
            return False
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]
            return True
        except Exception:
            return False

    def _teardown_native_windows_drop_target(self) -> None:
        if self._native_windows_drop_hwnd is not None:
            revoke_drop_target(self._native_windows_drop_hwnd)
        self._native_windows_drop_hwnd = None
        self._native_windows_drop_target = None
        self._native_windows_drop_target_com = None

    def _queue_native_drop_paths(self, paths: list[str]) -> None:
        self._native_drop_queue.put(("paths", list(paths or [])))

    def _queue_native_drop_error(self, details: str) -> None:
        self._native_drop_queue.put(("error", str(details)))

    def _poll_native_drop_queue(self) -> None:
        while True:
            try:
                event_type, payload = self._native_drop_queue.get_nowait()
            except queue.Empty:
                break

            if event_type == "paths":
                self._handle_native_windows_drop_paths(payload)
            elif event_type == "error":
                self._show_native_drop_error(str(payload))

        if self.root.winfo_exists():
            self.root.after(50, self._poll_native_drop_queue)

    def _show_native_drop_error(self, details: str) -> None:
        messagebox.showerror(
            APP_NAME,
            "The app couldn't read the dropped Outlook item.\n\n"
            f"{details}",
        )

    def _handle_native_windows_drop_paths(self, paths: list[str]) -> None:
        existing = [str(Path(p)) for p in (paths or []) if p and Path(p).exists()]
        if not existing:
            messagebox.showwarning(
                APP_NAME,
                "The dropped Outlook item could not be materialised into a local file.",
            )
            return
        self._handle_dropped_paths(existing)

    def _handle_dropped_paths(self, raw_paths: Iterable[str]) -> None:
        dropped_files: list[str] = []
        unsupported: list[Path] = []

        for item in raw_paths:
            path = Path(item)
            if path.is_dir():
                supported, folder_unsupported = get_supported_files_from_folder(path)
                dropped_files.extend(str(p) for p in supported)
                unsupported.extend(folder_unsupported)
            else:
                dropped_files.append(str(path))

        self._append_paths(dropped_files, precomputed_unsupported=unsupported)

    def _get_outlook_drop_temp_dir(self) -> Path:
        if self._outlook_drop_temp_dir_obj is None:
            self._outlook_drop_temp_dir_obj = tempfile.TemporaryDirectory(prefix="print_assist_outlook_drop_")
        target_dir = Path(self._outlook_drop_temp_dir_obj.name) / "outlook_attachments"
        target_dir.mkdir(parents=True, exist_ok=True)
        return target_dir

    def _cleanup_outlook_drop_temp_dir(self) -> None:
        if self._outlook_drop_temp_dir_obj is not None:
            self._outlook_drop_temp_dir_obj.cleanup()
            self._outlook_drop_temp_dir_obj = None

    def _materialise_outlook_virtual_attachments(
        self,
        names: list[str],
        payloads: Iterable[bytes | None],
    ) -> list[str]:
        target_dir = self._get_outlook_drop_temp_dir()
        result: list[str] = []

        for index, (name, payload) in enumerate(zip(names or [], payloads), start=1):
            if payload is None:
                continue
            payload_bytes = bytes(payload)
            fallback_name = f"attachment_{index}.bin"
            safe_name = safe_outlook_attachment_name(name, fallback_name)
            if payload_bytes.startswith(b"%PDF") and Path(safe_name).suffix.lower() != ".pdf":
                stem = Path(safe_name).stem or f"attachment_{index}"
                safe_name = f"{stem}.pdf"
            out_path = unique_file_path(target_dir / safe_name)
            out_path.write_bytes(payload_bytes)
            result.append(str(out_path))

        return result

    def _on_drop(self, event: tk.Event) -> None:
        raw = self.root.tk.splitlist(event.data)
        self._handle_dropped_paths(raw)

    def add_files(self) -> None:
        types = [("PDF", "*.pdf"), ("Images", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"), ("Word documents", "*.doc *.docx"), ("Excel workbooks", "*.xls *.xlsx *.xlsm *.xlsb"), ("Outlook messages", "*.msg"), ("ZIP archives", "*.zip"), ("All files", "*.*")]
        selected = filedialog.askopenfilenames(title="Select files", filetypes=types)
        self._append_paths(selected)

    def add_folder(self) -> None:
        selected_folder = filedialog.askdirectory(title="Select folder")
        if not selected_folder:
            return

        supported, unsupported = get_supported_files_from_folder(Path(selected_folder))
        added = self._append_paths(
            [str(path) for path in supported],
            precomputed_unsupported=unsupported,
        )

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
        added = self._append_paths(
            [str(path) for path in supported],
            precomputed_unsupported=unsupported,
        )

        if added:
            self.status_var.set(f"Added {added} file(s) from client folder.")
        elif supported:
            self.status_var.set("No new files were added from client folder.")
        else:
            self.status_var.set("No supported files found in the selected client folder.")

    def _expand_outlook_message_paths(
        self,
        paths: list[Path],
        max_nested_depth: int = 5,
    ) -> tuple[list[Path], list[Path], list[str]]:
        expanded: list[Path] = []
        unsupported: list[Path] = []
        warnings: list[str] = []
        visited_messages: set[Path] = set()
        attachment_dir = self._get_outlook_drop_temp_dir() / "message_attachments"

        def add_path(path: Path, depth: int) -> None:
            expanded.append(path)
            if path.suffix.lower() != ".msg":
                return

            try:
                message_key = path.resolve()
            except OSError:
                message_key = path
            if message_key in visited_messages:
                return
            visited_messages.add(message_key)

            try:
                attachments, attachment_warnings = extract_msg_attachments(path, attachment_dir)
                warnings.extend(f"{path.name}: {warning}" for warning in attachment_warnings)
            except Exception as exc:
                warnings.append(f"Could not include attachments from '{path.name}': {exc}")
                return

            for attachment_path in attachments:
                suffix = attachment_path.suffix.lower()
                if suffix not in PRINTABLE_EXTENSIONS:
                    unsupported.append(attachment_path)
                    continue
                if suffix == ".msg":
                    if depth >= max_nested_depth:
                        expanded.append(attachment_path)
                        warnings.append(
                            f"Nested attachment limit reached for '{attachment_path.name}'; "
                            "its own attachments were not expanded."
                        )
                    else:
                        add_path(attachment_path, depth + 1)
                else:
                    expanded.append(attachment_path)

        for path in paths:
            add_path(path, 0)

        return expanded, unsupported, warnings

    def _append_paths(
        self,
        raw_paths: tuple[str, ...] | list[str],
        precomputed_unsupported: list[Path] | None = None,
    ) -> int:
        supported, unsupported = filter_supported_files(raw_paths)
        if precomputed_unsupported:
            unsupported = list(unsupported) + precomputed_unsupported

        new_supported = [path for path in supported if path not in self.files]
        expanded, message_unsupported, message_warnings = self._expand_outlook_message_paths(new_supported)
        unsupported = list(unsupported) + message_unsupported

        added = 0
        for p in expanded:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, str(p))
                added += 1

        if added and self.output_path is None:
            self.output_path = default_output_path(new_supported or self.files)
            self.output_var.set(f"Output: {self.output_path}")

        warning_sections: list[str] = []
        if unsupported:
            display_unsupported = unsupported[:20]
            warning_msg = "Unsupported or non-printable attachments skipped:\n" + "\n".join(
                u.name for u in display_unsupported
            )
            remaining = len(unsupported) - len(display_unsupported)
            if remaining > 0:
                warning_msg += f"\n...and {remaining} more unsupported file(s)."
            warning_sections.append(warning_msg)
        if message_warnings:
            warning_sections.append("Outlook message warnings:\n" + "\n".join(message_warnings))
        if warning_sections:
            messagebox.showwarning(APP_NAME, "\n\n".join(warning_sections))

        self._update_file_count()
        self.status_var.set(f"{len(self.files)} file(s) selected.")
        return added

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
            "Rename + Extract ZIP",
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
                processed, warnings, preview_pdf, file_manifest = payload
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
                    file_manifest=file_manifest,
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

        zip_files = [p for p in self.files if p.suffix.lower() in ZIP_EXTENSIONS]
        if zip_files:
            messagebox.showerror(APP_NAME, "ZIP files cannot be previewed or printed directly. Use Rename + Extract ZIP first, then select printable files.")
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
                file_manifest: list[dict[str, object]] = []

                def manifest_callback(entry: dict[str, object]) -> None:
                    file_manifest.append(entry)

                processed, warnings = build_combined_pdf(
                    files_to_process,
                    preview_pdf,
                    progress_callback=progress_callback,
                    manifest_callback=manifest_callback,
                )
                if self._preview_queue is not None:
                    self._preview_queue.put(("done", (processed, warnings, preview_pdf, file_manifest)))
            except Exception as exc:
                if self._preview_queue is not None:
                    self._preview_queue.put(("error", f"Failed to create preview PDF:\n{exc}"))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self._poll_preview_queue)

    def rename_zip_contents(self) -> None:
        selected_indices = list(self.listbox.curselection())
        selected_zips = [self.files[i] for i in selected_indices if self.files[i].suffix.lower() in ZIP_EXTENSIONS]
        zip_files = selected_zips or [p for p in self.files if p.suffix.lower() in ZIP_EXTENSIONS]

        if not zip_files:
            selected = filedialog.askopenfilenames(title="Select ZIP files", filetypes=[("ZIP archives", "*.zip")])
            zip_files = [Path(p) for p in selected]

        if not zip_files:
            self.status_var.set("No ZIP files selected.")
            return

        outputs: list[Path] = []
        errors: list[str] = []
        warnings: list[str] = []

        if len(zip_files) == 1:
            source_zip = zip_files[0]
            selected_output = filedialog.askdirectory(
                title="Choose extraction folder",
                initialdir=str(source_zip.parent),
            )
            if not selected_output:
                self.status_var.set("ZIP rename and extract cancelled.")
                return
            output_folder = unique_folder_path(default_extracted_folder_path(source_zip, Path(selected_output)))
            try:
                outputs.append(rename_and_extract_zip_contents(source_zip, output_folder))
            except ZipExtractionWarning as exc:
                outputs.append(exc.output_folder)
                warnings.append(f"{source_zip.name}: {exc}")
            except Exception as exc:
                errors.append(f"{source_zip.name}: {exc}")
        else:
            selected_output = filedialog.askdirectory(
                title="Choose parent extraction folder",
                initialdir=str(zip_files[0].parent),
            )
            if not selected_output:
                self.status_var.set("ZIP rename and extract cancelled.")
                return
            parent_output = Path(selected_output)
            for source_zip in zip_files:
                output_folder = unique_folder_path(default_extracted_folder_path(source_zip, parent_output))
                try:
                    outputs.append(rename_and_extract_zip_contents(source_zip, output_folder))
                except ZipExtractionWarning as exc:
                    outputs.append(exc.output_folder)
                    warnings.append(f"{source_zip.name}: {exc}")
                except Exception as exc:
                    errors.append(f"{source_zip.name}: {exc}")

        if outputs:
            output_text = "\n".join(str(p) for p in outputs)
            self.status_var.set(f"Renamed and extracted {len(outputs)} ZIP file(s).")
            messagebox.showinfo(APP_NAME, f"Renamed ZIP contents extracted to:\n{output_text}")

        if warnings:
            self.status_var.set("ZIP contents extracted with warnings.")
            messagebox.showwarning(APP_NAME, "ZIP rename and extract warnings:\n" + "\n".join(warnings))

        if errors:
            self.status_var.set("Some ZIP files could not be renamed and extracted.")
            messagebox.showerror(APP_NAME, "ZIP rename and extract errors:\n" + "\n".join(errors))

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

    def _on_close(self) -> None:
        self._teardown_native_windows_drop_target()
        if self._windows_ole_initialized:
            ole_uninitialize()
            self._windows_ole_initialized = False
        self._cleanup_preview_temp_dir()
        self._cleanup_outlook_drop_temp_dir()
        self.root.destroy()


def run(root: tk.Tk | None = None) -> None:
    root = root if root is not None else tk.Tk()
    app = PrintAssistApp(root)
    _ = app
    root.mainloop()
