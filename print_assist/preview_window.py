from __future__ import annotations

import copy
import os
import shutil
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import fitz
from PIL import Image, ImageTk

from .mouse_scroll import bind_mouse_scroll

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
EDIT_MARGIN = 24
MIN_CROP_SCREEN_PIXELS = 12


def _fit_rect(
    src_width: float,
    src_height: float,
    page_width: float,
    page_height: float,
    margin: float,
) -> fitz.Rect:
    usable_width = page_width - (2 * margin)
    usable_height = page_height - (2 * margin)
    scale = min(usable_width / src_width, usable_height / src_height)
    draw_width = src_width * scale
    draw_height = src_height * scale
    x0 = (page_width - draw_width) / 2
    y0 = (page_height - draw_height) / 2
    return fitz.Rect(x0, y0, x0 + draw_width, y0 + draw_height)


def _update_manifest_after_page_removal(
    file_manifest: list[dict[str, object]],
    target_entry_index: int,
    kept_end_page: int,
    removed_count: int,
) -> list[dict[str, object]]:
    updated = copy.deepcopy(file_manifest)
    target = updated[target_entry_index]
    start_page = target.get("output_start_page")
    if not isinstance(start_page, int):
        raise ValueError("The email source page range is unavailable.")

    target["output_end_page"] = kept_end_page
    target["output_page_count"] = kept_end_page - start_page + 1

    for entry in updated[target_entry_index + 1 :]:
        entry_start = entry.get("output_start_page")
        entry_end = entry.get("output_end_page")
        if isinstance(entry_start, int):
            entry["output_start_page"] = entry_start - removed_count
        if isinstance(entry_end, int):
            entry["output_end_page"] = entry_end - removed_count
    return updated


def _update_manifest_after_single_page_deletion(
    file_manifest: list[dict[str, object]],
    target_entry_index: int,
    deleted_page: int,
) -> list[dict[str, object]]:
    updated = copy.deepcopy(file_manifest)
    target = updated[target_entry_index]
    start_page = target.get("output_start_page")
    end_page = target.get("output_end_page")
    if (
        not isinstance(start_page, int)
        or not isinstance(end_page, int)
        or not start_page <= deleted_page <= end_page
    ):
        raise ValueError("The source page range is unavailable.")

    remaining_count = end_page - start_page
    if remaining_count > 0:
        target["output_end_page"] = end_page - 1
        target["output_page_count"] = remaining_count
    else:
        # Keep the source in File Summary while showing that it no longer
        # contributes a page to the edited preview.
        target["output_start_page"] = None
        target["output_end_page"] = None
        target["output_page_count"] = 0

    for entry in updated[target_entry_index + 1 :]:
        entry_start = entry.get("output_start_page")
        entry_end = entry.get("output_end_page")
        if isinstance(entry_start, int):
            entry["output_start_page"] = entry_start - 1
        if isinstance(entry_end, int):
            entry["output_end_page"] = entry_end - 1
    return updated


def _format_preview_page_status(
    page_number: int,
    total_pages: int,
    file_manifest: list[dict[str, object]],
) -> str:
    prefix = f"Overall page {page_number} of {total_pages}"
    for entry in file_manifest:
        start_page = entry.get("output_start_page")
        end_page = entry.get("output_end_page")
        if (
            not isinstance(start_page, int)
            or not isinstance(end_page, int)
            or not start_page <= page_number <= end_page
        ):
            continue

        source_name = entry.get("source_name")
        file_page = page_number - start_page + 1
        file_page_count = end_page - start_page + 1
        if isinstance(source_name, str) and source_name:
            return (
                f"{prefix} — {source_name} — "
                f"File page {file_page} of {file_page_count}"
            )
        return f"{prefix} — File page {file_page} of {file_page_count}"
    return prefix


class PreviewWindow:
    def __init__(
        self,
        parent: tk.Tk,
        preview_pdf_path: Path,
        output_path: Path,
        open_pdf_callback,
        on_status_change,
        on_close_callback,
        file_manifest: list[dict[str, object]] | None = None,
    ) -> None:
        self.parent = parent
        self.preview_pdf_path = preview_pdf_path
        self.output_path = output_path
        self.open_pdf_callback = open_pdf_callback
        self.on_status_change = on_status_change
        self.on_close_callback = on_close_callback
        self._closed = False
        self.file_manifest = file_manifest or []

        self.zoom = 1.0
        self.page_index = 0
        self.saved_final = False
        self._photo: ImageTk.PhotoImage | None = None
        self._edit_mode: str | None = None
        self._selection_start: tuple[float, float] | None = None
        self._selection_item: int | None = None
        self._undo_stack: list[tuple[bytes, list[dict[str, object]], int]] = []

        self.doc = fitz.open(str(self.preview_pdf_path))
        self._original_pdf_bytes = self.doc.tobytes(garbage=4, deflate=True)
        self._original_manifest = copy.deepcopy(self.file_manifest)

        self.window = tk.Toplevel(parent)
        self.window.title("Print Assist Preview")
        self.window.geometry("1000x760")
        self.window.minsize(640, 480)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        self._build_ui()
        self._render_page()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.window, padding=8)
        top.pack(fill=tk.X)

        self.page_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.page_var).pack(side=tk.LEFT)

        ttk.Label(top, text=f"Output: {self.output_path}").pack(side=tk.RIGHT)

        self.canvas_frame = ttk.Frame(self.window)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.canvas = tk.Canvas(self.canvas_frame, background="#202020", highlightthickness=0)
        self.v_scroll = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        self.canvas_frame.rowconfigure(0, weight=1)
        self.canvas_frame.columnconfigure(0, weight=1)
        bind_mouse_scroll(self.canvas)
        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)

        buttons = ttk.Frame(self.window, padding=(8, 0, 8, 8))
        buttons.pack(fill=tk.X)

        edit_buttons = ttk.Frame(buttons)
        edit_buttons.pack(fill=tk.X)
        output_buttons = ttk.Frame(buttons)
        output_buttons.pack(fill=tk.X)

        self.prev_btn = ttk.Button(edit_buttons, text="Previous Page", command=self.prev_page)
        self.next_btn = ttk.Button(edit_buttons, text="Next Page", command=self.next_page)
        self.crop_image_btn = ttk.Button(edit_buttons, text="Crop Image", command=self.start_image_crop)
        self.trim_email_btn = ttk.Button(
            edit_buttons,
            text="Trim Email Below Line",
            command=self.start_email_trim,
        )
        self.delete_page_btn = ttk.Button(
            edit_buttons,
            text="Delete Current Page",
            command=self.delete_current_page,
        )
        self.undo_btn = ttk.Button(edit_buttons, text="Undo Edit", command=self.undo_edit)
        self.reset_btn = ttk.Button(edit_buttons, text="Reset Edits", command=self.reset_edits)
        ttk.Button(edit_buttons, text="Zoom Out", command=self.zoom_out).pack(
            side=tk.LEFT, padx=4, pady=4
        )
        ttk.Button(edit_buttons, text="Zoom In", command=self.zoom_in).pack(
            side=tk.LEFT, padx=4, pady=4
        )
        self.prev_btn.pack(side=tk.LEFT, padx=4, pady=4)
        self.next_btn.pack(side=tk.LEFT, padx=4, pady=4)
        self.delete_page_btn.pack(side=tk.LEFT, padx=4, pady=4)
        self.trim_email_btn.pack(side=tk.LEFT, padx=4, pady=4)
        self.crop_image_btn.pack(side=tk.LEFT, padx=4, pady=4)
        self.undo_btn.pack(side=tk.LEFT, padx=4, pady=4)
        self.reset_btn.pack(side=tk.LEFT, padx=4, pady=4)
        ttk.Button(output_buttons, text="Save Final PDF", command=self.save_final_pdf).pack(side=tk.RIGHT, padx=4, pady=4)
        ttk.Button(output_buttons, text="File Summary", command=self.open_file_summary).pack(side=tk.RIGHT, padx=4, pady=4)
        ttk.Button(output_buttons, text="Open Preview Externally", command=self.open_preview_externally).pack(side=tk.RIGHT, padx=4, pady=4)
        ttk.Button(output_buttons, text="Close", command=self.close).pack(side=tk.RIGHT, padx=4, pady=4)

        self.edit_var = tk.StringVar(value="")
        ttk.Label(self.window, textvariable=self.edit_var, padding=(12, 0, 12, 8)).pack(anchor="w")

    def _render_page(self) -> None:
        self._cancel_edit_mode()
        page = self.doc[self.page_index]
        matrix = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self._photo = ImageTk.PhotoImage(image)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self._photo)
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))

        total = len(self.doc)
        page_number = self.page_index + 1
        self.page_var.set(
            _format_preview_page_status(page_number, total, self.file_manifest)
        )
        self.prev_btn.configure(state=tk.NORMAL if self.page_index > 0 else tk.DISABLED)
        self.next_btn.configure(state=tk.NORMAL if self.page_index < total - 1 else tk.DISABLED)
        self._update_edit_controls()

    def _get_manifest_entry_for_page(
        self,
        page_number: int,
    ) -> tuple[int, dict[str, object]] | None:
        for index, entry in enumerate(self.file_manifest):
            start_page = entry.get("output_start_page")
            end_page = entry.get("output_end_page")
            if (
                isinstance(start_page, int)
                and isinstance(end_page, int)
                and start_page <= page_number <= end_page
            ):
                return index, entry
        return None

    def _get_source_name_for_page(self, page_number: int) -> str | None:
        match = self._get_manifest_entry_for_page(page_number)
        if match is not None:
            source_name = match[1].get("source_name")
            if isinstance(source_name, str):
                return source_name
        return None

    def _update_edit_controls(self) -> None:
        match = self._get_manifest_entry_for_page(self.page_index + 1)
        source_extension = ""
        if match is not None:
            extension = match[1].get("source_extension")
            if isinstance(extension, str):
                source_extension = extension.lower()

        self.crop_image_btn.configure(
            state=tk.NORMAL if source_extension in IMAGE_EXTENSIONS else tk.DISABLED
        )
        self.trim_email_btn.configure(
            state=tk.NORMAL if source_extension == ".msg" else tk.DISABLED
        )
        self.delete_page_btn.configure(
            state=tk.NORMAL if len(self.doc) > 1 else tk.DISABLED
        )
        self.undo_btn.configure(state=tk.NORMAL if self._undo_stack else tk.DISABLED)
        has_edits = bool(self._undo_stack) or self.file_manifest != self._original_manifest
        self.reset_btn.configure(state=tk.NORMAL if has_edits else tk.DISABLED)

    def _cancel_edit_mode(self) -> None:
        self._edit_mode = None
        self._selection_start = None
        if self._selection_item is not None:
            self.canvas.delete(self._selection_item)
            self._selection_item = None
        self.canvas.configure(cursor="")
        if hasattr(self, "edit_var"):
            self.edit_var.set("")

    def start_image_crop(self) -> None:
        self._cancel_edit_mode()
        self._edit_mode = "crop_image"
        self.canvas.configure(cursor="crosshair")
        self.edit_var.set("Drag a rectangle around the part of the image you want to keep.")

    def start_email_trim(self) -> None:
        self._cancel_edit_mode()
        self._edit_mode = "trim_email"
        self.canvas.configure(cursor="crosshair")
        self.edit_var.set(
            "Click on the email where it should end. Content below the line and later pages "
            "from this email will be removed."
        )

    def _canvas_point(self, event: tk.Event) -> tuple[float, float]:
        page = self.doc[self.page_index]
        max_x = page.rect.width * self.zoom
        max_y = page.rect.height * self.zoom
        x = min(max(self.canvas.canvasx(event.x), 0), max_x)
        y = min(max(self.canvas.canvasy(event.y), 0), max_y)
        return x, y

    def _on_canvas_press(self, event: tk.Event) -> None:
        if self._edit_mode is None:
            return
        x, y = self._canvas_point(event)
        self._selection_start = (x, y)
        if self._edit_mode == "crop_image":
            self._selection_item = self.canvas.create_rectangle(
                x,
                y,
                x,
                y,
                outline="#ffcc00",
                width=3,
                dash=(8, 4),
            )
        else:
            page_width = self.doc[self.page_index].rect.width * self.zoom
            self._selection_item = self.canvas.create_line(
                0,
                y,
                page_width,
                y,
                fill="#ff4444",
                width=3,
                dash=(8, 4),
            )

    def _on_canvas_drag(self, event: tk.Event) -> None:
        if self._selection_start is None or self._selection_item is None:
            return
        x, y = self._canvas_point(event)
        start_x, start_y = self._selection_start
        if self._edit_mode == "crop_image":
            self.canvas.coords(self._selection_item, start_x, start_y, x, y)
        elif self._edit_mode == "trim_email":
            page_width = self.doc[self.page_index].rect.width * self.zoom
            self.canvas.coords(self._selection_item, 0, y, page_width, y)

    def _on_canvas_release(self, event: tk.Event) -> None:
        if self._selection_start is None or self._edit_mode is None:
            return
        end_x, end_y = self._canvas_point(event)
        start_x, start_y = self._selection_start
        mode = self._edit_mode
        self._cancel_edit_mode()

        if mode == "crop_image":
            if (
                abs(end_x - start_x) < MIN_CROP_SCREEN_PIXELS
                or abs(end_y - start_y) < MIN_CROP_SCREEN_PIXELS
            ):
                messagebox.showinfo(
                    "Print Assist",
                    "Drag a larger rectangle around the image area to keep.",
                    parent=self.window,
                )
                return
            clip = fitz.Rect(
                min(start_x, end_x) / self.zoom,
                min(start_y, end_y) / self.zoom,
                max(start_x, end_x) / self.zoom,
                max(start_y, end_y) / self.zoom,
            )
            try:
                self._apply_image_crop(clip)
            except Exception as exc:
                messagebox.showerror(
                    "Print Assist",
                    f"Could not crop this image page:\n{exc}",
                    parent=self.window,
                )
        else:
            try:
                self._apply_email_trim(end_y / self.zoom)
            except Exception as exc:
                messagebox.showerror(
                    "Print Assist",
                    f"Could not trim this email page:\n{exc}",
                    parent=self.window,
                )

    def _push_undo_state(self) -> None:
        self._undo_stack.append(
            (
                self.doc.tobytes(garbage=4, deflate=True),
                copy.deepcopy(self.file_manifest),
                self.page_index,
            )
        )

    def _replace_preview_document(
        self,
        new_doc: fitz.Document,
        new_manifest: list[dict[str, object]],
        new_page_index: int,
    ) -> None:
        replacement_path = self.preview_pdf_path.with_name("preview.replacement.pdf")
        try:
            if replacement_path.exists():
                replacement_path.unlink()
            new_doc.set_metadata(self.doc.metadata)
            new_doc.save(replacement_path, garbage=4, deflate=True)
        finally:
            new_doc.close()

        self.doc.close()
        os.replace(replacement_path, self.preview_pdf_path)
        self.doc = fitz.open(str(self.preview_pdf_path))
        self.file_manifest = new_manifest
        self.page_index = min(max(new_page_index, 0), len(self.doc) - 1)
        self.saved_final = False
        self._render_page()

    @staticmethod
    def _insert_page_range(
        destination: fitz.Document,
        source: fitz.Document,
        first_page: int,
        last_page: int,
    ) -> None:
        if first_page <= last_page:
            destination.insert_pdf(source, from_page=first_page, to_page=last_page)

    def _apply_image_crop(self, clip: fitz.Rect) -> None:
        page_rect = self.doc[self.page_index].rect
        clip = clip & page_rect
        if clip.is_empty or clip.width < 1 or clip.height < 1:
            return

        self._push_undo_state()
        source = fitz.open(stream=self._undo_stack[-1][0], filetype="pdf")
        edited = fitz.open()
        try:
            self._insert_page_range(edited, source, 0, self.page_index - 1)
            output_page = edited.new_page(width=page_rect.width, height=page_rect.height)
            target = _fit_rect(
                clip.width,
                clip.height,
                page_rect.width,
                page_rect.height,
                EDIT_MARGIN,
            )
            output_page.show_pdf_page(target, source, self.page_index, clip=clip)
            self._insert_page_range(edited, source, self.page_index + 1, len(source) - 1)
            self._replace_preview_document(
                edited,
                copy.deepcopy(self.file_manifest),
                self.page_index,
            )
        except Exception:
            edited.close()
            self._undo_stack.pop()
            raise
        finally:
            source.close()
        self.on_status_change("Image cropped in preview")

    def _apply_email_trim(self, cut_y: float) -> None:
        match = self._get_manifest_entry_for_page(self.page_index + 1)
        if match is None or str(match[1].get("source_extension", "")).lower() != ".msg":
            return
        entry_index, entry = match
        end_page = entry.get("output_end_page")
        if not isinstance(end_page, int):
            return

        page_rect = self.doc[self.page_index].rect
        cut_y = min(max(cut_y, 36), page_rect.height)
        source_end_index = end_page - 1
        removed_count = max(0, source_end_index - self.page_index)

        self._push_undo_state()
        source = fitz.open(stream=self._undo_stack[-1][0], filetype="pdf")
        edited = fitz.open()
        try:
            self._insert_page_range(edited, source, 0, self.page_index - 1)
            output_page = edited.new_page(width=page_rect.width, height=page_rect.height)
            clip = fitz.Rect(0, 0, page_rect.width, cut_y)
            output_page.show_pdf_page(clip, source, self.page_index, clip=clip)
            self._insert_page_range(edited, source, source_end_index + 1, len(source) - 1)

            new_manifest = _update_manifest_after_page_removal(
                self.file_manifest,
                entry_index,
                self.page_index + 1,
                removed_count,
            )
            self._replace_preview_document(edited, new_manifest, self.page_index)
        except Exception:
            edited.close()
            self._undo_stack.pop()
            raise
        finally:
            source.close()
        self.on_status_change("Email thread trimmed in preview")

    def delete_current_page(self) -> None:
        if len(self.doc) <= 1:
            messagebox.showinfo(
                "Print Assist",
                "The only remaining page cannot be deleted.",
                parent=self.window,
            )
            return

        try:
            self._apply_page_deletion()
        except Exception as exc:
            messagebox.showerror(
                "Print Assist",
                f"Could not delete this preview page:\n{exc}",
                parent=self.window,
            )

    def _apply_page_deletion(self) -> None:
        match = self._get_manifest_entry_for_page(self.page_index + 1)
        if match is None:
            raise ValueError("The source for this page could not be identified.")
        entry_index, _entry = match
        deleted_page = self.page_index + 1

        self._push_undo_state()
        source = fitz.open(stream=self._undo_stack[-1][0], filetype="pdf")
        edited = fitz.open()
        try:
            self._insert_page_range(edited, source, 0, self.page_index - 1)
            self._insert_page_range(edited, source, self.page_index + 1, len(source) - 1)
            new_manifest = _update_manifest_after_single_page_deletion(
                self.file_manifest,
                entry_index,
                deleted_page,
            )
            new_page_index = min(self.page_index, len(source) - 2)
            self._replace_preview_document(edited, new_manifest, new_page_index)
        except Exception:
            edited.close()
            self._undo_stack.pop()
            raise
        finally:
            source.close()
        self.on_status_change("Current page deleted from preview only")

    def _restore_pdf_state(
        self,
        pdf_bytes: bytes,
        manifest: list[dict[str, object]],
        page_index: int,
    ) -> None:
        restored = fitz.open(stream=pdf_bytes, filetype="pdf")
        self._replace_preview_document(restored, copy.deepcopy(manifest), page_index)

    def undo_edit(self) -> None:
        if not self._undo_stack:
            return
        pdf_bytes, manifest, page_index = self._undo_stack.pop()
        self._restore_pdf_state(pdf_bytes, manifest, page_index)
        self.on_status_change("Last preview edit undone")

    def reset_edits(self) -> None:
        if not self._undo_stack and self.file_manifest == self._original_manifest:
            return
        reset = messagebox.askyesno(
            "Print Assist",
            "Reset all image crops, email trims, and deleted pages?",
            parent=self.window,
        )
        if not reset:
            return
        self._undo_stack.clear()
        original_doc = fitz.open(stream=self._original_pdf_bytes, filetype="pdf")
        try:
            restored_page_index = min(self.page_index, original_doc.page_count - 1)
        finally:
            original_doc.close()
        self._restore_pdf_state(
            self._original_pdf_bytes,
            self._original_manifest,
            restored_page_index,
        )
        self.on_status_change("All preview edits reset")

    def open_file_summary(self) -> None:
        summary_window = tk.Toplevel(self.window)
        summary_window.title("File Summary")
        summary_window.geometry("980x340")
        summary_window.minsize(700, 220)

        container = ttk.Frame(summary_window, padding=8)
        container.pack(fill=tk.BOTH, expand=True)

        columns = ("order", "name", "type", "range", "count", "path")
        tree = ttk.Treeview(container, columns=columns, show="headings")
        tree.heading("order", text="Order")
        tree.heading("name", text="File name")
        tree.heading("type", text="Type")
        tree.heading("range", text="Page range")
        tree.heading("count", text="Page count")
        tree.heading("path", text="Full path")

        tree.column("order", width=70, minwidth=60, anchor=tk.CENTER)
        tree.column("name", width=200, minwidth=160)
        tree.column("type", width=90, minwidth=80, anchor=tk.CENTER)
        tree.column("range", width=110, minwidth=100, anchor=tk.CENTER)
        tree.column("count", width=90, minwidth=80, anchor=tk.CENTER)
        tree.column("path", width=400, minwidth=260)

        y_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
        x_scroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        bind_mouse_scroll(tree)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        for idx, entry in enumerate(self.file_manifest, start=1):
            start_page = entry.get("output_start_page", "")
            end_page = entry.get("output_end_page", "")
            page_range = (
                f"{start_page}-{end_page}"
                if isinstance(start_page, int) and isinstance(end_page, int)
                else "Deleted from preview"
            )
            tree.insert(
                "",
                tk.END,
                values=(
                    idx,
                    entry.get("source_name", ""),
                    entry.get("source_extension", ""),
                    page_range,
                    entry.get("output_page_count", ""),
                    entry.get("source_path", ""),
                ),
            )

    def prev_page(self) -> None:
        if self.page_index > 0:
            self.page_index -= 1
            self._render_page()

    def next_page(self) -> None:
        if self.page_index < len(self.doc) - 1:
            self.page_index += 1
            self._render_page()

    def zoom_out(self) -> None:
        self.zoom = max(0.4, self.zoom - 0.2)
        self._render_page()

    def zoom_in(self) -> None:
        self.zoom = min(3.0, self.zoom + 0.2)
        self._render_page()

    def save_final_pdf(self) -> None:
        if self.output_path.exists():
            overwrite = messagebox.askyesno("Print Assist", f"Output file exists. Overwrite?\n{self.output_path}", parent=self.window)
            if not overwrite:
                return

        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.preview_pdf_path, self.output_path)
        self.saved_final = True
        self.on_status_change("Final PDF saved")

        open_now = messagebox.askyesno("Print Assist", f"Final PDF saved:\n{self.output_path}\n\nOpen now?", parent=self.window)
        if open_now:
            self.open_pdf_callback(self.output_path)

    def open_preview_externally(self) -> None:
        self.open_pdf_callback(self.preview_pdf_path)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self.doc.close()
        self.window.destroy()
        self.on_close_callback()
