import tkinter as tk

from print_assist.app import run
from print_assist.app_icon import configure_windows_app_identity


if __name__ == "__main__":
    configure_windows_app_identity()
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()
    run(root)
