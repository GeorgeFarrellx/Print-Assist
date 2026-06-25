from __future__ import annotations

import unittest
from types import SimpleNamespace

from print_assist.mouse_scroll import _scroll_from_wheel, _wheel_scroll_units


class FakeScrollableWidget:
    def __init__(self) -> None:
        self.vertical_calls: list[tuple[int, str]] = []
        self.horizontal_calls: list[tuple[int, str]] = []

    def yview_scroll(self, units: int, mode: str) -> None:
        self.vertical_calls.append((units, mode))

    def xview_scroll(self, units: int, mode: str) -> None:
        self.horizontal_calls.append((units, mode))


class MouseScrollTests(unittest.TestCase):
    def test_windows_wheel_delta_is_converted_to_scroll_units(self) -> None:
        self.assertEqual(_wheel_scroll_units(SimpleNamespace(delta=120)), -1)
        self.assertEqual(_wheel_scroll_units(SimpleNamespace(delta=-240)), 2)

    def test_linux_wheel_buttons_are_converted_to_scroll_units(self) -> None:
        self.assertEqual(_wheel_scroll_units(SimpleNamespace(num=4, delta=0)), -1)
        self.assertEqual(_wheel_scroll_units(SimpleNamespace(num=5, delta=0)), 1)

    def test_wheel_scrolls_vertically_over_widget_content(self) -> None:
        widget = FakeScrollableWidget()

        result = _scroll_from_wheel(
            widget,
            SimpleNamespace(delta=-120, state=0),
        )

        self.assertEqual(result, "break")
        self.assertEqual(widget.vertical_calls, [(1, "units")])
        self.assertEqual(widget.horizontal_calls, [])

    def test_shift_wheel_scrolls_horizontally(self) -> None:
        widget = FakeScrollableWidget()

        result = _scroll_from_wheel(
            widget,
            SimpleNamespace(delta=-120, state=1),
        )

        self.assertEqual(result, "break")
        self.assertEqual(widget.vertical_calls, [])
        self.assertEqual(widget.horizontal_calls, [(1, "units")])


if __name__ == "__main__":
    unittest.main()
