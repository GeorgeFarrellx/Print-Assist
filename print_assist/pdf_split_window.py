from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import fitz
from PIL import Image, ImageTk

from .app_icon import configure_window_icon
from .mouse_scroll import bind_mouse_scroll
from .pdf_splitter import (
    PdfInfo,
    PdfPart,
    PdfSplitResult,
    build_split_plan,
    export_split_pdf,
    inspect_pdf,
    parse_page_ranges,
    part_filename,
    source_signature,
)


class PdfSplitWindow:
    def __init__(
        self,
        parent: tk.Misc,
        source_path: Path,
        on_export: Callable[[PdfSplitResult], None],
        on_close: Callable[[], None],
    ) -> None:
        self.source_path = source_path
        self.on_export = on_export
        self.on_close = on_close
        self.info: PdfInfo | None = None
        self.parts: tuple[PdfPart, ...] = ()
        self.doc: fitz.Document | None = None
        self.page_index = 0
        self.zoom: float | None = None
        self.fit_width = True
        self._render_scale = 1.0
        self._photo: ImageTk.PhotoImage | None = None
        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._closed = False
        self._exporting = False
        self._render_after: str | None = None
        self._poll_after: str | None = None
        self._last_result: PdfSplitResult | None = None
        self._controls: list[ttk.Widget] = []

        self.window = tk.Toplevel(parent)
        self.window.title("Print Assist - Split PDF")
        self.window.geometry("1100x800")
        self.window.minsize(900, 680)
        self.window.transient(parent)
        configure_window_icon(self.window)
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        self.window.bind("<Escape>", lambda _event: self.close())

        self.ranges_var = tk.StringVar(master=self.window, value="")
        self.prefix_var = tk.StringVar(master=self.window, value=source_path.stem)
        self.folder_var = tk.StringVar(master=self.window, value=str(source_path.parent))
        self.status_var = tk.StringVar(master=self.window, value="Opening PDF...")
        self.summary_var = tk.StringVar(master=self.window, value="Reading PDF...")
        self.page_var = tk.StringVar(master=self.window, value="1")
        self.page_label_var = tk.StringVar(master=self.window, value="")
        self.progress_var = tk.DoubleVar(master=self.window, value=0)

        self._build_ui()
        self._set_controls_enabled(False)
        self.prefix_var.trace_add("write", lambda *_args: self._update_plan())
        self.window.grab_set()
        threading.Thread(target=self._inspect_worker, daemon=True).start()
        self._poll_after = self.window.after(100, self._poll_queue)

    def _button(self, parent: tk.Misc, text: str, command) -> ttk.Button:
        button = ttk.Button(parent, text=text, command=command)
        self._controls.append(button)
        return button

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.window, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        header = ttk.Frame(frame)
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Split PDF", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(header, text=self.source_path.name, wraplength=850).pack(anchor="w", pady=(2, 6))

        ttk.Label(header, text="Preview a page and mark where a new PDF begins, or enter the page ranges below.", wraplength=1000).pack(anchor="w")
        inputs = ttk.Frame(header)
        inputs.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(inputs, text="Page ranges (one PDF per range):").pack(side=tk.LEFT)
        self.ranges_entry = ttk.Entry(inputs, textvariable=self.ranges_var)
        self.ranges_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        self.ranges_entry.bind("<Return>", lambda _event: self._update_plan())
        self._button(inputs, "Apply Ranges", self._update_plan).pack(side=tk.LEFT)
        ttk.Label(header, text="Example: 1-5, 6-10, 11-20. Include every page once, in order.").pack(anchor="w", pady=(4, 0))
        ttk.Label(header, textvariable=self.summary_var).pack(anchor="w", pady=(4, 6))

        panes = ttk.Panedwindow(frame, orient=tk.HORIZONTAL)
        panes.grid(row=1, column=0, sticky="nsew")
        files_frame = ttk.Frame(panes)
        preview_frame = ttk.Frame(panes)
        panes.add(files_frame, weight=1)
        panes.add(preview_frame, weight=1)
        ttk.Label(files_frame, text="Files to create", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tree_frame = ttk.Frame(files_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 0), padx=(0, 8))
        self.tree = ttk.Treeview(
            tree_frame, columns=("pages", "count", "filename"), show="headings", selectmode="browse"
        )
        for column, label, width in (("pages", "Pages", 65), ("count", "Count", 50), ("filename", "Filename", 380)):
            self.tree.heading(column, text=label)
            self.tree.column(column, width=width, minwidth=width, stretch=column == "filename")
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._select_part)
        bind_mouse_scroll(self.tree)

        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)
        preview_heading = ttk.Label(preview_frame, text="Original page preview", font=("Segoe UI", 10, "bold"))
        preview_heading.grid(row=0, column=0, sticky="w")
        canvas_frame = ttk.Frame(preview_frame)
        canvas_frame.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.canvas = tk.Canvas(canvas_frame, background="#eeeeee", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        preview_y_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        preview_y_scroll.grid(row=0, column=1, sticky="ns")
        preview_x_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        preview_x_scroll.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=preview_y_scroll.set, xscrollcommand=preview_x_scroll.set)
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)
        bind_mouse_scroll(self.canvas)
        self.canvas.bind("<Configure>", self._schedule_render)
        self.preview_controls = ttk.Frame(preview_frame)
        self.preview_controls.grid(row=2, column=0, sticky="ew")
        nav = ttk.Frame(self.preview_controls)
        nav.pack(fill=tk.X, pady=(4, 0))
        self._button(nav, "Previous", lambda: self._show_page(self.page_index - 1)).pack(side=tk.LEFT)
        self._button(nav, "Next", lambda: self._show_page(self.page_index + 1)).pack(side=tk.LEFT, padx=4)
        ttk.Label(nav, text="Page:").pack(side=tk.LEFT, padx=(6, 0))
        self.page_entry = ttk.Entry(nav, textvariable=self.page_var, width=5)
        self.page_entry.pack(side=tk.LEFT, padx=4)
        self.page_entry.bind("<Return>", lambda _event: self._go_to_page())
        self._button(nav, "Go", self._go_to_page).pack(side=tk.LEFT)
        zoom_row = ttk.Frame(self.preview_controls)
        zoom_row.pack(fill=tk.X, pady=(4, 0))
        self._button(zoom_row, "Zoom Out", lambda: self._change_zoom(1 / 1.25)).pack(side=tk.LEFT)
        self._button(zoom_row, "Zoom In", lambda: self._change_zoom(1.25)).pack(side=tk.LEFT, padx=4)
        self._button(zoom_row, "Fit Width", self._fit_width).pack(side=tk.LEFT, padx=(0, 4))
        self._button(zoom_row, "Fit Page", self._fit_page).pack(side=tk.LEFT)
        ttk.Label(self.preview_controls, textvariable=self.page_label_var, wraplength=480).pack(anchor="w", pady=4)
        edits = ttk.Frame(self.preview_controls)
        edits.pack(fill=tk.X)
        self._button(edits, "Split before this page", self._add_boundary).pack(side=tk.LEFT)
        self._button(edits, "Remove split here", self._remove_boundary).pack(side=tk.LEFT, padx=4)

        self.footer = ttk.Frame(frame)
        self.footer.grid(row=2, column=0, sticky="ew")
        output = ttk.Frame(self.footer)
        output.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(output, text="Filename prefix:").grid(row=0, column=0, sticky="w")
        self.prefix_entry = ttk.Entry(output, textvariable=self.prefix_var)
        self.prefix_entry.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(output, text="Save inside:").grid(row=1, column=0, sticky="w", pady=6)
        self.folder_entry = ttk.Entry(output, textvariable=self.folder_var)
        self.folder_entry.grid(row=1, column=1, sticky="ew", padx=6)
        self._button(output, "Choose Folder", self._choose_folder).grid(row=1, column=2)
        output.columnconfigure(1, weight=1)
        ttk.Label(self.footer, text="Creates a new folder. Every original page is kept once, in its original order.").pack(anchor="w")
        ttk.Progressbar(self.footer, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=(6, 0))
        ttk.Label(self.footer, textvariable=self.status_var, wraplength=1000).pack(anchor="w", pady=4)
        actions = ttk.Frame(self.footer)
        actions.pack(fill=tk.X)
        ttk.Button(actions, text="Close", command=self.close).pack(side=tk.RIGHT)
        self.export_button = self._button(actions, "Export Split PDFs", self._export)
        self.export_button.pack(side=tk.RIGHT, padx=6)
        self.open_folder_button = ttk.Button(actions, text="Open Export Folder", command=self._open_folder, state=tk.DISABLED)
        self.open_folder_button.pack(side=tk.LEFT)

        self.window.update_idletasks()
        self.window.minsize(
            max(900, self.preview_controls.winfo_reqwidth() + 540),
            max(540, header.winfo_reqheight() + self.footer.winfo_reqheight()
                + self.preview_controls.winfo_reqheight() + preview_heading.winfo_reqheight() + 100),
        )

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        for widget in self._controls:
            widget.configure(state=state)
        for entry in (self.prefix_entry, self.folder_entry, self.page_entry, self.ranges_entry):
            entry.configure(state=state)

    def _inspect_worker(self) -> None:
        try:
            info = inspect_pdf(self.source_path)
            self._queue.put(("inspected", info))
        except Exception as exc:
            self._queue.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        self._poll_after = None
        if self._closed:
            return
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "export_progress":
                    current, total = payload
                    self.progress_var.set(current * 100 / total)
                    self.status_var.set(f"Exporting file {current} of {total}...")
                elif kind == "inspected":
                    try:
                        self.info = payload
                        self.doc = fitz.open(self.source_path)
                        if source_signature(self.source_path) != self.info.source_signature:
                            raise ValueError("The source PDF changed. Close and reopen Split PDF.")
                        self.ranges_var.set(f"1-{self.info.page_count}" if self.info.page_count > 1 else "1")
                        self._set_controls_enabled(True)
                        self._update_plan()
                        self._show_page(0)
                    except Exception as exc:
                        self.status_var.set(str(exc))
                        self.info = None
                        self._set_controls_enabled(False)
                elif kind == "exported":
                    self._exporting = False
                    self._last_result = payload
                    self._set_controls_enabled(True)
                    self._update_plan()
                    self.open_folder_button.configure(state=tk.NORMAL)
                    self.status_var.set(f"Saved {len(payload.files)} PDFs in: {payload.folder}")
                    self.on_export(payload)
                elif kind == "error":
                    self._exporting = False
                    self._set_controls_enabled(self.info is not None)
                    self.status_var.set(f"Could not complete split: {payload}")
                    messagebox.showerror("Split PDF", str(payload), parent=self.window)
        except queue.Empty:
            pass
        self._poll_after = self.window.after(100, self._poll_queue)

    def _update_plan(self) -> bool:
        if self.info is None or self._exporting or self._closed:
            return False
        self._set_controls_enabled(True)
        try:
            parts = parse_page_ranges(self.ranges_var.get(), self.info.page_count)
        except ValueError as exc:
            self.parts = ()
            self.tree.delete(*self.tree.get_children())
            self.summary_var.set("Correct the split settings to continue.")
            self.status_var.set(str(exc))
            self.export_button.configure(state=tk.DISABLED)
            return False
        self.parts = parts
        self.tree.delete(*self.tree.get_children())
        for index, part in enumerate(parts, start=1):
            self.tree.insert("", tk.END, iid=str(index - 1), values=(part.page_range, part.page_count, part_filename(self.prefix_var.get(), index, part)))
        self.summary_var.set(f"{len(parts)} PDF file(s)  |  {self.info.page_count} of {self.info.page_count} pages included")
        if self.info.page_count == 1:
            self.status_var.set("This PDF has only one page. Choose a PDF with at least two pages to split.")
        else:
            self.status_var.set("Mark a split on the preview or enter your page ranges." if len(parts) == 1 else "Review your chosen ranges, then export.")
        self.export_button.configure(state=tk.NORMAL if len(parts) > 1 else tk.DISABLED)
        self._update_page_label()
        return True

    def _select_part(self, _event: tk.Event | None = None) -> None:
        selected = self.tree.selection()
        if selected and self.parts:
            self._show_page(self.parts[int(selected[0])].start_page - 1)

    def _show_page(self, index: int) -> None:
        if self.doc is None or self._exporting or not 0 <= index < self.doc.page_count:
            return
        self.page_index = index
        self.page_var.set(str(index + 1))
        self._update_page_label()
        self._render_page()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def _update_page_label(self) -> None:
        if self.info is None:
            return
        label = f"Page {self.page_index + 1} of {self.info.page_count}"
        for index, part in enumerate(self.parts, start=1):
            if part.start_page <= self.page_index + 1 <= part.end_page:
                label += f"  |  File {index}, pages {part.page_range}"
                break
        self.page_label_var.set(label)

    def _go_to_page(self) -> None:
        try:
            page = int(self.page_var.get())
            if self.info is None or not 1 <= page <= self.info.page_count:
                raise ValueError
        except ValueError:
            self.status_var.set("Enter a page number within this PDF.")
            return
        self._show_page(page - 1)

    def _schedule_render(self, _event: tk.Event | None = None) -> None:
        if self._render_after is not None:
            self.window.after_cancel(self._render_after)
        self._render_after = self.window.after(100, self._render_page)

    def _render_page(self) -> None:
        if self._render_after is not None:
            self.window.after_cancel(self._render_after)
        self._render_after = None
        if self.doc is None or self._closed:
            return
        try:
            page = self.doc[self.page_index]
            width, height = self.canvas.winfo_width(), self.canvas.winfo_height()
            if width < 20 or height < 20:
                return
            fit_scale = min((width - 16) / page.rect.width, 2.0)
            if not self.fit_width:
                fit_scale = min(fit_scale, (height - 16) / page.rect.height)
            scale = self.zoom or fit_scale
            self._render_scale = scale
            pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            self._photo = ImageTk.PhotoImage(Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples), master=self.window)
            self.canvas.delete("all")
            draw_width, draw_height = max(width, pixmap.width + 16), max(height, pixmap.height + 16)
            self.canvas.create_image(draw_width / 2, draw_height / 2, anchor="center", image=self._photo)
            self.canvas.configure(scrollregion=(0, 0, draw_width, draw_height))
        except Exception as exc:
            self.status_var.set(f"Cannot preview this page: {exc}")

    def _change_zoom(self, factor: float) -> None:
        self.zoom = min(3.0, max(0.2, self._render_scale * factor))
        self._render_page()

    def _fit_page(self) -> None:
        self.zoom = None
        self.fit_width = False
        self._render_page()

    def _fit_width(self) -> None:
        self.zoom = None
        self.fit_width = True
        self._render_page()

    def _edit_boundary(self, add: bool) -> None:
        if self.info is None or self._exporting or not self._update_plan():
            return
        page = self.page_index + 1
        if page == 1:
            self.status_var.set("Page 1 always starts the first file.")
            return
        starts = {part.start_page for part in self.parts}
        if not add and page not in starts:
            self.status_var.set("This page does not start a new file. Preview the first page of a file to remove its split.")
            return
        if add:
            starts.add(page)
        else:
            starts.discard(page)
        parts = build_split_plan(self.info.page_count, starts)
        self.ranges_var.set(", ".join(part.page_range for part in parts))
        self._update_plan()

    def _add_boundary(self) -> None:
        self._edit_boundary(True)

    def _remove_boundary(self) -> None:
        self._edit_boundary(False)

    def _choose_folder(self) -> None:
        selected = filedialog.askdirectory(parent=self.window, title="Save split PDFs inside this folder", initialdir=self.folder_var.get())
        if selected:
            self.folder_var.set(selected)

    def _export(self) -> None:
        if not self._update_plan() or len(self.parts) < 2:
            return
        destination = Path(self.folder_var.get().strip()).expanduser()
        if not destination.is_dir():
            self.status_var.set("Choose an existing folder to save the split PDFs inside.")
            return
        parts, prefix = self.parts, self.prefix_var.get()
        signature = self.info.source_signature
        self._exporting = True
        self._set_controls_enabled(False)
        self.progress_var.set(0)
        self.status_var.set("Exporting split PDFs...")

        def worker() -> None:
            try:
                result = export_split_pdf(
                    self.source_path, parts, destination, prefix,
                    lambda current, total: self._queue.put(("export_progress", (current, total))),
                    expected_signature=signature,
                )
                self._queue.put(("exported", result))
            except Exception as exc:
                self._queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _open_folder(self) -> None:
        if self._last_result is not None:
            try:
                os.startfile(self._last_result.folder)
            except OSError as exc:
                self.status_var.set(f"Could not open the folder: {exc}")

    def close(self) -> bool:
        if self._closed:
            return True
        if self._exporting:
            self.status_var.set("Please wait for the export to finish before closing.")
            return False
        self._closed = True
        if self._render_after is not None:
            self.window.after_cancel(self._render_after)
        if self._poll_after is not None:
            self.window.after_cancel(self._poll_after)
        if self.doc is not None:
            self.doc.close()
        if self.window.grab_current() == self.window:
            self.window.grab_release()
        self.window.destroy()
        self.on_close()
        return True
