from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path


ASSET_DIR = Path(__file__).resolve().parent / "assets"
WINDOW_ICON_SIZES = (256, 128, 64, 48, 32, 24, 16)


def icon_asset_path(filename: str) -> Path:
    return ASSET_DIR / filename


def configure_window_icon(root: tk.Misc) -> bool:
    """Apply the bundled icon set to a Tk root without preventing startup."""
    applied = False
    images: list[tk.PhotoImage] = []

    for size in WINDOW_ICON_SIZES:
        image_path = icon_asset_path(f"print-assist-{size}.png")
        if not image_path.is_file():
            continue
        try:
            images.append(tk.PhotoImage(master=root, file=str(image_path)))
        except tk.TclError:
            continue

    if images:
        try:
            root.iconphoto(True, *images)
            # Tk does not retain Python references to icon images.
            setattr(root, "_print_assist_icon_images", tuple(images))
            applied = True
        except tk.TclError:
            pass

    # On Windows, iconphoto and iconbitmap override one another. Keep the
    # exact-size transparent PNG set when it loaded successfully, and use the
    # ICO only as a compatibility fallback.
    if sys.platform == "win32" and not applied:
        ico_path = icon_asset_path("print-assist.ico")
        if ico_path.is_file():
            try:
                root.iconbitmap(str(ico_path))
                applied = True
            except tk.TclError:
                pass

    return applied
