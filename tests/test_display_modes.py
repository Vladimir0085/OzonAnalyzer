from __future__ import annotations

import inspect
import unittest

from ozon_app.display_modes import fitted_window_size, resolve_ui_scale
from ozon_app.resizable_layout import ResizableOZPriceAnalyzerApp, _minimum_upper_height
from ozon_app.ui import OZPriceAnalyzerApp


class DisplayModeTests(unittest.TestCase):
    def test_auto_scale_is_compact_on_laptop_heights(self) -> None:
        self.assertEqual(resolve_ui_scale("auto", 768), 0.8)
        self.assertEqual(resolve_ui_scale("auto", 900), 0.9)
        self.assertEqual(resolve_ui_scale("auto", 1080), 1.0)

    def test_explicit_scale_overrides_screen_height(self) -> None:
        self.assertEqual(resolve_ui_scale("100%", 768), 1.0)
        self.assertEqual(resolve_ui_scale("0.9", 1080), 0.9)
        self.assertEqual(resolve_ui_scale("80%", 1080), 0.8)

    def test_window_size_never_uses_desktop_sized_minimum_on_small_screen(self) -> None:
        self.assertEqual(fitted_window_size(1366, 768), (1318, 688, 1180, 608))
        self.assertEqual(fitted_window_size(1920, 1080), (1540, 920, 1180, 720))

    def test_protected_settings_rows_cannot_be_collapsed_by_splitter(self) -> None:
        self.assertEqual(_minimum_upper_height(265, []), 265)
        self.assertEqual(_minimum_upper_height(212, [24, 220, 92]), 348)
        self.assertEqual(_minimum_upper_height(400, [24, 220, 92]), 400)

    def test_overview_requests_an_equal_adaptive_table_split(self) -> None:
        base_source = inspect.getsource(OZPriceAnalyzerApp._build_overview_tab)
        layout_source = inspect.getsource(ResizableOZPriceAnalyzerApp._build_overview_tab)

        self.assertIn("table_fraction=0.5", layout_source)
        self.assertIn("CompactKpi.TLabel", base_source)
        self.assertIn("Compact.TCombobox", base_source)


if __name__ == "__main__":
    unittest.main()
