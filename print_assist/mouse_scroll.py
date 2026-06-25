from __future__ import annotations

import tkinter as tk
from typing import Any

SHIFT_MASK = 0x0001


def _wheel_scroll_units(event: tk.Event) -> int:
    """Return Tk scroll units for Windows/macOS and Linux wheel events."""
    button_number = getattr(event, "num", None)
    if button_number == 4:
        return -1
    if button_number == 5:
        return 1

    delta = int(getattr(event, "delta", 0) or 0)
    if delta == 0:
        return 0
    if abs(delta) >= 120:
        return -int(delta / 120)
    return -1 if delta > 0 else 1


def _scroll_from_wheel(widget: Any, event: tk.Event) -> str | None:
    units = _wheel_scroll_units(event)
    if units == 0:
        return None

    if int(getattr(event, "state", 0) or 0) & SHIFT_MASK:
        widget.xview_scroll(units, "units")
    else:
        widget.yview_scroll(units, "units")
    return "break"


def bind_mouse_scroll(widget: Any) -> None:
    """Make a Tk scrollable widget respond while the pointer is over its content."""

    def on_wheel(event: tk.Event) -> str | None:
        return _scroll_from_wheel(widget, event)

    widget.bind("<MouseWheel>", on_wheel)
    widget.bind("<Button-4>", on_wheel)
    widget.bind("<Button-5>", on_wheel)
