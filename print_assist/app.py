from __future__ import annotations

import os
import queue
import subprocess
import tempfile
import threading
import tkinter as tk
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES
except Exception:
    DND_FILES = None

from . import APP_NAME
from .app_icon import configure_window_icon
from .file_utils import (
    PRINTABLE_EXTENSIONS,
    ZIP_EXTENSIONS,
    default_output_path,
    filter_supported_files,
    get_supported_files_from_client_folder,
    get_supported_files_from_folder,
)
from .mouse_scroll import bind_mouse_scroll
from .outlook_message import extract_msg_details, safe_outlook_attachment_name, unique_file_path
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

SORT_MANUAL = "Manual order"
SORT_FILENAME = "File name (A–Z)"
SORT_EMAIL_DATE = "Email date (oldest first)"


def _pluralised_count(count: int, singular: str, plural: str | None = None) -> str:
    label = singular if count == 1 else (plural or f"{singular}s")
    return f"{count} {label}"


def _path_is_within(path: Path, directory: Path) -> bool:
    try:
        path.resolve().relative_to(directory.resolve())
        return True
    except (OSError, ValueError):
        return False


def format_file_selection_summary(
    files: Iterable[Path],
    outlook_temp_dir: Path | None = None,
) -> str:
    selected = list(files)
    email_count = sum(path.suffix.lower() == ".msg" for path in selected)
    attachment_count = 0

    if outlook_temp_dir is not None:
        attachment_count = sum(
            path.suffix.lower() != ".msg" and _path_is_within(path, outlook_temp_dir)
            for path in selected
        )

    if email_count == 0 and attachment_count == 0:
        return f"Selected files: {len(selected)}"

    other_count = len(selected) - email_count - attachment_count
    parts: list[str] = []
    if email_count:
        parts.append(_pluralised_count(email_count, "email"))
    if attachment_count:
        parts.append(_pluralised_count(attachment_count, "attachment"))
    if other_count:
        parts.append(_pluralised_count(other_count, "other file"))
    return f"Selected: {' + '.join(parts)}"


def reorder_grouped_files(
    files: Iterable[Path],
    parents: dict[Path, Path | None],
    selected: set[Path],
    direction: int,
) -> list[Path]:
    ordered_files = list(files)
    file_set = set(ordered_files)
    children: dict[Path | None, list[Path]] = {}
    for path in ordered_files:
        parent = parents.get(path)
        children.setdefault(parent if parent in file_set else None, []).append(path)

    changed = False
    for siblings in children.values():
        indices = range(1, len(siblings)) if direction < 0 else range(len(siblings) - 2, -1, -1)
        for index in indices:
            adjacent_index = index + direction
            if siblings[index] in selected and siblings[adjacent_index] not in selected:
                siblings[adjacent_index], siblings[index] = siblings[index], siblings[adjacent_index]
                changed = True

    if not changed:
        return ordered_files

    reordered: list[Path] = []

    def append_branch(parent: Path | None) -> None:
        for path in children.get(parent, []):
            reordered.append(path)
            append_branch(path)

    append_branch(None)
    return reordered


def sort_grouped_files(
    files: Iterable[Path],
    parents: dict[Path, Path | None],
    sort_mode: str,
    email_datetimes: dict[Path, datetime | None] | None = None,
) -> list[Path]:
    manual_files = list(files)
    if sort_mode == SORT_MANUAL:
        return manual_files

    manual_index = {path: index for index, path in enumerate(manual_files)}
    file_set = set(manual_files)
    children: dict[Path | None, list[Path]] = {}
    for path in manual_files:
        parent = parents.get(path)
        children.setdefault(parent if parent in file_set else None, []).append(path)

    if sort_mode == SORT_FILENAME:
        children.get(None, []).sort(
            key=lambda path: (path.name.casefold(), manual_index[path])
        )
    elif sort_mode == SORT_EMAIL_DATE:
        dates = email_datetimes or {}

        def email_date_key(path: Path) -> tuple[int, float, int]:
            value = dates.get(path)
            if value is None:
                return (1, 0.0, manual_index[path])
            try:
                timestamp = value.timestamp()
            except (OSError, OverflowError, ValueError):
                return (1, 0.0, manual_index[path])
            return (0, timestamp, manual_index[path])

        children.get(None, []).sort(key=email_date_key)

    sorted_files: list[Path] = []

    def append_branch(parent: Path | None) -> None:
        for path in children.get(parent, []):
            sorted_files.append(path)
            append_branch(path)

    append_branch(None)
    return sorted_files


class PrintAssistApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        configure_window_icon(self.root)
        self.root.geometry("900x560")

        self.files: list[Path] = []
        self.manual_files: list[Path] = []
        self.file_parents: dict[Path, Path | None] = {}
        self.file_datetimes: dict[Path, datetime | None] = {}
        self._tree_item_paths: dict[str, Path] = {}
        self.output_path: Path | None = None

        self.status_var = tk.StringVar(value="Ready")
        self.output_var = tk.StringVar(value="No output file selected")
        self.progress_var = tk.DoubleVar(value=0)
        self.file_count_var = tk.StringVar(value="Selected files: 0")
        self.sort_var = tk.StringVar(value=SORT_MANUAL)
        self._preview_running = False
        self._preview_queue: queue.Queue[tuple[str, object]] | None = None
        self._preview_temp_dir_obj: tempfile.TemporaryDirectory[str] | None = None
        self._preview_view: PreviewWindow | None = None
        self._main_window_geometry = "900x560"
        self._main_window_minsize = self.root.minsize()
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
        self.main_view = ttk.Frame(self.root)
        self.main_view.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(self.main_view, text=APP_NAME, font=("Segoe UI", 18, "bold"))
        title.pack(pady=(12, 8))

        frame = ttk.Frame(self.main_view, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        list_label = ttk.Label(frame, text="Drop files, folders, or Outlook email messages here:")
        list_label.pack(anchor="w")

        sort_frame = ttk.Frame(frame)
        sort_frame.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(sort_frame, text="Sort by:").pack(side=tk.LEFT)
        self.sort_combo = ttk.Combobox(
            sort_frame,
            textvariable=self.sort_var,
            values=(SORT_MANUAL, SORT_FILENAME, SORT_EMAIL_DATE),
            state="readonly",
            width=28,
        )
        self.sort_combo.pack(side=tk.LEFT, padx=(6, 0))
        self.sort_combo.bind("<<ComboboxSelected>>", self._on_sort_changed)

        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 10))

        self.file_tree = ttk.Treeview(
            list_frame,
            columns=("date_time", "folder"),
            selectmode="extended",
            show=("tree", "headings"),
            height=18,
        )
        self.file_tree.heading("#0", text="File")
        self.file_tree.heading("date_time", text="Email date & time")
        self.file_tree.heading("folder", text="Folder")
        self.file_tree.column("#0", width=300, minwidth=180, stretch=True)
        self.file_tree.column("date_time", width=145, minwidth=130, stretch=False)
        self.file_tree.column("folder", width=370, minwidth=180, stretch=True)
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.file_tree.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.file_tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        bind_mouse_scroll(self.file_tree)
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
            hwnd = int(self.file_tree.winfo_id())
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
        expanded, _, _, unsupported, warnings = PrintAssistApp._expand_outlook_message_entries(
            self,
            paths,
            max_nested_depth=max_nested_depth,
        )
        return expanded, unsupported, warnings

    def _expand_outlook_message_entries(
        self,
        paths: list[Path],
        max_nested_depth: int = 5,
    ) -> tuple[
        list[Path],
        dict[Path, Path | None],
        dict[Path, datetime | None],
        list[Path],
        list[str],
    ]:
        expanded: list[Path] = []
        parents: dict[Path, Path | None] = {}
        message_datetimes: dict[Path, datetime | None] = {}
        unsupported: list[Path] = []
        warnings: list[str] = []
        visited_messages: set[Path] = set()
        attachment_dir = self._get_outlook_drop_temp_dir() / "message_attachments"

        def add_path(path: Path, depth: int, parent_email: Path | None = None) -> None:
            expanded.append(path)
            parents[path] = parent_email
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
                attachments, attachment_warnings, message_datetime = extract_msg_details(
                    path,
                    attachment_dir,
                )
                message_datetimes[path] = message_datetime
                warnings.extend(f"{path.name}: {warning}" for warning in attachment_warnings)
            except Exception as exc:
                message_datetimes[path] = None
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
                        parents[attachment_path] = path
                        message_datetimes[attachment_path] = None
                        warnings.append(
                            f"Nested attachment limit reached for '{attachment_path.name}'; "
                            "its own attachments were not expanded."
                        )
                    else:
                        add_path(attachment_path, depth + 1, path)
                else:
                    expanded.append(attachment_path)
                    parents[attachment_path] = path

        for path in paths:
            add_path(path, 0)

        return expanded, parents, message_datetimes, unsupported, warnings

    def _append_paths(
        self,
        raw_paths: tuple[str, ...] | list[str],
        precomputed_unsupported: list[Path] | None = None,
    ) -> int:
        supported, unsupported = filter_supported_files(raw_paths)
        if precomputed_unsupported:
            unsupported = list(unsupported) + precomputed_unsupported

        new_supported = [path for path in supported if path not in self.manual_files]
        expanded, expanded_parents, expanded_datetimes, message_unsupported, message_warnings = (
            self._expand_outlook_message_entries(new_supported)
        )
        unsupported = list(unsupported) + message_unsupported

        added = 0
        for p in expanded:
            if p not in self.manual_files:
                self.manual_files.append(p)
                self.file_parents[p] = expanded_parents.get(p)
                self.file_datetimes[p] = expanded_datetimes.get(p)
                added += 1
        if added:
            self._apply_sort()

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
        outlook_temp_dir = None
        if self._outlook_drop_temp_dir_obj is not None:
            outlook_temp_dir = Path(self._outlook_drop_temp_dir_obj.name)
        self.file_count_var.set(format_file_selection_summary(self.files, outlook_temp_dir))

    def remove_selected(self) -> None:
        selected = set(self._selected_tree_paths())
        if not selected:
            self.status_var.set("No files selected.")
            return

        removed = set(selected)
        changed = True
        while changed:
            changed = False
            for path in self.manual_files:
                if path not in removed and self.file_parents.get(path) in removed:
                    removed.add(path)
                    changed = True

        self.manual_files = [path for path in self.manual_files if path not in removed]
        for path in removed:
            self.file_parents.pop(path, None)
            self.file_datetimes.pop(path, None)
        self._apply_sort()
        self._update_file_count()
        self.status_var.set(f"{len(removed)} file(s) removed.")

    def move_up(self) -> None:
        self._move_selected_groups(-1)

    def move_down(self) -> None:
        self._move_selected_groups(1)

    def _on_sort_changed(self, event: tk.Event | None = None) -> None:
        _ = event
        self._apply_sort()
        self.status_var.set(f"Sorted by {self.sort_var.get().lower()}.")

    def _apply_sort(self, selected_paths: Iterable[Path] | None = None) -> None:
        self.files = sort_grouped_files(
            self.manual_files,
            self.file_parents,
            self.sort_var.get(),
            self.file_datetimes,
        )
        self._refresh_file_tree(selected_paths)

    def _format_file_datetime(self, path: Path) -> str:
        value = self.file_datetimes.get(path)
        if value is None:
            return ""
        return value.strftime("%d/%m/%Y %H:%M")

    def _selected_tree_paths(self) -> list[Path]:
        return [
            self._tree_item_paths[item_id]
            for item_id in self.file_tree.selection()
            if item_id in self._tree_item_paths
        ]

    def _refresh_file_tree(self, selected_paths: Iterable[Path] | None = None) -> None:
        if selected_paths is None:
            selected_paths = self._selected_tree_paths()
        selected = set(selected_paths)
        open_paths = {
            path
            for item_id, path in self._tree_item_paths.items()
            if self.file_tree.exists(item_id) and bool(self.file_tree.item(item_id, "open"))
        }
        previously_rendered_paths = set(self._tree_item_paths.values())

        root_items = self.file_tree.get_children()
        if root_items:
            self.file_tree.delete(*root_items)
        self._tree_item_paths.clear()
        path_items: dict[Path, str] = {}
        selected_items: list[str] = []
        parent_paths = {parent for parent in self.file_parents.values() if parent is not None}

        for index, path in enumerate(self.files):
            parent_path = self.file_parents.get(path)
            parent_item = path_items.get(parent_path, "")
            item_id = f"file_{index}"
            has_children = path in parent_paths
            self.file_tree.insert(
                parent_item,
                tk.END,
                iid=item_id,
                text=path.name,
                values=(self._format_file_datetime(path), str(path.parent)),
                open=path in open_paths or (has_children and path not in previously_rendered_paths),
            )
            path_items[path] = item_id
            self._tree_item_paths[item_id] = path
            if path in selected:
                selected_items.append(item_id)

        if selected_items:
            self.file_tree.selection_set(selected_items)
            self.file_tree.see(selected_items[0])

    def _move_selected_groups(self, direction: int) -> None:
        selected = set(self._selected_tree_paths())
        if not selected:
            self.status_var.set("No files selected.")
            return

        if self.sort_var.get() != SORT_MANUAL:
            self.sort_var.set(SORT_MANUAL)

        reordered = reorder_grouped_files(self.manual_files, self.file_parents, selected, direction)
        if reordered == self.manual_files:
            self._apply_sort(selected)
            return

        self.manual_files = reordered
        self._apply_sort(selected)

    def clear_files(self) -> None:
        self.files.clear()
        self.manual_files.clear()
        self.file_parents.clear()
        self.file_datetimes.clear()
        self._refresh_file_tree(())
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
        self.sort_combo.configure(state="readonly" if enabled else tk.DISABLED)

    def _cleanup_preview_temp_dir(self) -> None:
        if self._preview_temp_dir_obj is not None:
            self._preview_temp_dir_obj.cleanup()
            self._preview_temp_dir_obj = None

    def _show_preview(
        self,
        preview_pdf: Path,
        file_manifest: list[dict[str, object]],
    ) -> None:
        self._main_window_geometry = self.root.geometry()
        self._main_window_minsize = self.root.minsize()
        self.main_view.pack_forget()
        self.root.title(f"{APP_NAME} - Preview")
        self.root.geometry("1000x760")
        self.root.minsize(640, 480)
        try:
            self._preview_view = PreviewWindow(
                parent=self.root,
                preview_pdf_path=preview_pdf,
                output_path=self.output_path,
                open_pdf_callback=self.open_pdf,
                on_status_change=self.status_var.set,
                on_close_callback=self._restore_main_view,
                file_manifest=file_manifest,
            )
        except Exception:
            self._restore_main_view()
            raise

    def _restore_main_view(self) -> None:
        self._preview_view = None
        self._cleanup_preview_temp_dir()
        self.root.title(APP_NAME)
        self.root.minsize(*self._main_window_minsize)
        self.root.geometry(self._main_window_geometry)
        self.main_view.pack(fill=tk.BOTH, expand=True)
        self.status_var.set("Preview closed")

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
                try:
                    self._show_preview(preview_pdf, file_manifest)
                except Exception as exc:
                    self.progress_var.set(0)
                    self.status_var.set("Error")
                    messagebox.showerror(APP_NAME, f"Failed to open preview:\n{exc}")
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
        selected_zips = [
            path for path in self._selected_tree_paths() if path.suffix.lower() in ZIP_EXTENSIONS
        ]
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
        if self._preview_view is not None:
            self._preview_view.close()
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
