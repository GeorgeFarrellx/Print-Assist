from __future__ import annotations

import unittest
from unittest.mock import patch

from print_assist import app_icon


class _FakeRoot:
    def __init__(self) -> None:
        self.iconphoto_args: tuple[object, ...] | None = None
        self.iconbitmap_path: str | None = None

    def iconphoto(self, *args: object) -> None:
        self.iconphoto_args = args

    def iconbitmap(self, path: str) -> None:
        self.iconbitmap_path = path


class AppIconTests(unittest.TestCase):
    def test_supplied_icon_assets_are_present(self) -> None:
        expected = {
            "print-assist.ico",
            "print-assist.svg",
            *(f"print-assist-{size}.png" for size in app_icon.WINDOW_ICON_SIZES),
        }

        self.assertEqual(
            expected,
            {path.name for path in app_icon.ASSET_DIR.iterdir() if path.is_file()},
        )

    def test_configure_window_icon_prefers_exact_size_pngs_on_windows(self) -> None:
        root = _FakeRoot()

        with (
            patch.object(app_icon.sys, "platform", "win32"),
            patch.object(
                app_icon.tk,
                "PhotoImage",
                side_effect=lambda *, master, file: (master, file),
            ),
        ):
            applied = app_icon.configure_window_icon(root)

        self.assertTrue(applied)
        self.assertIsNotNone(root.iconphoto_args)
        self.assertEqual(True, root.iconphoto_args[0])
        self.assertEqual(len(app_icon.WINDOW_ICON_SIZES), len(root.iconphoto_args) - 1)
        self.assertIsNone(root.iconbitmap_path)
        self.assertEqual(
            len(app_icon.WINDOW_ICON_SIZES),
            len(root._print_assist_icon_images),
        )

    def test_configure_window_icon_uses_ico_when_pngs_cannot_load(self) -> None:
        root = _FakeRoot()

        with (
            patch.object(app_icon.sys, "platform", "win32"),
            patch.object(app_icon.tk, "PhotoImage", side_effect=app_icon.tk.TclError),
        ):
            applied = app_icon.configure_window_icon(root)

        self.assertTrue(applied)
        self.assertIsNone(root.iconphoto_args)
        self.assertTrue(root.iconbitmap_path.endswith("print-assist.ico"))


if __name__ == "__main__":
    unittest.main()
