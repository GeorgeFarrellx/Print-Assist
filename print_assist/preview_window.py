from __future__ import annotations

import os
import shutil
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import fitz
from PIL import Image, ImageTk


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

        self.doc = fitz.open(str(self.preview_pdf_path))

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

        buttons = ttk.Frame(self.window, padding=(8, 0, 8, 8))
        buttons.pack(fill=tk.X)

        self.prev_btn = ttk.Button(buttons, text="Previous Page", command=self.prev_page)
        self.next_btn = ttk.Button(buttons, text="Next Page", command=self.next_page)
        ttk.Button(buttons, text="Zoom Out", command=self.zoom_out).pack(side=tk.LEFT, padx=4, pady=4)
        ttk.Button(buttons, text="Zoom In", command=self.zoom_in).pack(side=tk.LEFT, padx=4, pady=4)
        self.prev_btn.pack(side=tk.LEFT, padx=4, pady=4)
        self.next_btn.pack(side=tk.LEFT, padx=4, pady=4)
        ttk.Button(buttons, text="Save Final PDF", command=self.save_final_pdf).pack(side=tk.RIGHT, padx=4, pady=4)
        ttk.Button(buttons, text="File Summary", command=self.open_file_summary).pack(side=tk.RIGHT, padx=4, pady=4)
        ttk.Button(buttons, text="Print Preview PDF", command=self.print_preview_pdf).pack(side=tk.RIGHT, padx=4, pady=4)
        ttk.Button(buttons, text="Open Preview Externally", command=self.open_preview_externally).pack(side=tk.RIGHT, padx=4, pady=4)
        ttk.Button(buttons, text="Close", command=self.close).pack(side=tk.RIGHT, padx=4, pady=4)

    def _render_page(self) -> None:
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
        source_name = self._get_source_name_for_page(page_number)
        if source_name:
            self.page_var.set(f"Page {page_number} of {total} — Source: {source_name}")
        else:
            self.page_var.set(f"Page {page_number} of {total}")
        self.prev_btn.configure(state=tk.NORMAL if self.page_index > 0 else tk.DISABLED)
        self.next_btn.configure(state=tk.NORMAL if self.page_index < total - 1 else tk.DISABLED)


    def _get_source_name_for_page(self, page_number: int) -> str | None:
        for entry in self.file_manifest:
            start_page = entry.get("output_start_page")
            end_page = entry.get("output_end_page")
            if isinstance(start_page, int) and isinstance(end_page, int) and start_page <= page_number <= end_page:
                source_name = entry.get("source_name")
                if isinstance(source_name, str):
                    return source_name
        return None

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
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        for idx, entry in enumerate(self.file_manifest, start=1):
            start_page = entry.get("output_start_page", "")
            end_page = entry.get("output_end_page", "")
            page_range = f"{start_page}-{end_page}" if start_page != "" and end_page != "" else ""
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

    def print_preview_pdf(self) -> None:
        confirm = messagebox.askyesno(
            "Print Assist",
            "Print the preview PDF using your default Windows PDF printing setup?",
            parent=self.window,
        )
        if not confirm:
            return

        if sys.platform != "win32" or not hasattr(os, "startfile"):
            messagebox.showinfo(
                "Print Assist",
                "Direct printing is only supported through Windows/default PDF print handling.",
                parent=self.window,
            )
            return

        try:
            os.startfile(str(self.preview_pdf_path), "print")
            self.on_status_change("Print command sent")
        except Exception as exc:
            messagebox.showerror("Print Assist", f"Failed to send print command:\n{exc}", parent=self.window)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self.doc.close()
        self.window.destroy()
        self.on_close_callback()
