from __future__ import annotations

import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from tkinter import ttk
from unittest.mock import patch

from ozon_app.display_modes import LaptopFriendlyOZPriceAnalyzerApp
from ozon_app.report_totals import ReportTotalsOZPriceAnalyzerApp
from ozon_app.service import AppService
from ozon_app.ui_containers import ScrollableSummary, summary_scrollbars, wrapped_positions


class OverviewFilterLayoutTests(unittest.TestCase):
    def test_groups_wrap_without_clipping_at_laptop_and_desktop_widths(self) -> None:
        for width in (880, 1180, 1318, 1540, 1860):
            for scale in (0.8, 0.9, 1.0):
                with self.subTest(width=width, scale=scale):
                    sizes = [(round(w * scale), round(h * scale)) for w, h in
                             ((430, 36), (200, 32), (250, 32), (200, 32), (95, 36), (125, 24))]
                    positions, height = wrapped_positions(sizes, width)
                    for (x, y), (group_width, group_height) in zip(positions, sizes):
                        self.assertLessEqual(x + group_width, width)
                        self.assertLessEqual(y + group_height, height)

    def test_widening_window_unwraps_the_controls(self) -> None:
        sizes = [(400, 36), (250, 32), (250, 32)]
        self.assertGreater(wrapped_positions(sizes, 700)[1], 36)
        self.assertEqual(wrapped_positions(sizes, 1200)[1], 36)

    def test_scrollbars_reserve_space_for_one_another(self) -> None:
        self.assertEqual(summary_scrollbars(1200, 320, 1100, 300, 15, 15), (False, False))
        self.assertEqual(summary_scrollbars(1200, 320, 1190, 330, 15, 15), (True, True))


class OverviewTkGeometryTests(unittest.TestCase):
    """Run with a real display on Windows CI (skipped on headless Linux)."""

    @classmethod
    def setUpClass(cls) -> None:
        try:
            probe = tk.Tk()
        except tk.TclError as exc:
            if "no display name" in str(exc) or "couldn't connect to display" in str(exc):
                raise unittest.SkipTest(f"Tk display is unavailable: {exc}") from exc
            raise
        probe.destroy()

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.callback_errors: list[tuple[object, ...]] = []
        error_patch = patch(
            "tkinter.Tk.report_callback_exception",
            side_effect=lambda *error: self.callback_errors.append(error),
        )
        error_patch.start()
        self.addCleanup(error_patch.stop)
        initialize_tk = tk.Tk.__init__

        def initialize_at_test_dpi(root, *args, **kwargs):
            initialize_tk(root, *args, **kwargs)
            root.tk.call("tk", "scaling", 96 / 72)

        # CI's virtual monitor must not limit the requested laptop/desktop widths.
        with patch("tkinter.Tk.__init__", initialize_at_test_dpi), patch.object(
            LaptopFriendlyOZPriceAnalyzerApp, "_maximize_on_small_screen", lambda self: None
        ):
            self.app = ReportTotalsOZPriceAnalyzerApp(AppService(Path(directory.name)))
        self.addCleanup(self.app.destroy)
        self.app.minsize(1, 1)
        self.app.maxsize(2200, 1400)
        self.app._base_tk_scaling = 96 / 72

    def settle(self) -> None:
        for _ in range(3):
            self.app.update_idletasks()
            self.app.update()
        self.assertEqual(self.callback_errors, [])

    def assert_inside(self, child, parent) -> None:
        self.assertTrue(child.winfo_viewable(), str(child))
        self.assertGreater(child.winfo_width(), 1)
        self.assertGreater(child.winfo_height(), 1)
        self.assertGreaterEqual(child.winfo_rootx(), parent.winfo_rootx())
        self.assertGreaterEqual(child.winfo_rooty(), parent.winfo_rooty())
        self.assertLessEqual(child.winfo_rootx() + child.winfo_width(),
                             parent.winfo_rootx() + parent.winfo_width())
        self.assertLessEqual(child.winfo_rooty() + child.winfo_height(),
                             parent.winfo_rooty() + parent.winfo_height())

    def assert_filters_visible(self) -> None:
        filters = self.app.overview_filters
        self.assert_inside(filters, self.app.overview_tab)
        self.assert_inside(filters, self.app._overview_splitter.bar)
        for group in filters.groups:
            self.assert_inside(group, filters)
            for child in group.winfo_children():
                if child.winfo_manager():
                    self.assert_inside(child, group)
        self.assertGreaterEqual(self.app.overview_tree.winfo_rooty(),
                                filters.winfo_rooty() + filters.winfo_height())

    def test_overview_filters_and_equal_table_split_across_screen_sizes(self) -> None:
        for scale in (0.8, 0.9, 1.0):
            self.app._apply_ui_scale(scale)
            for width, height in ((1180, 608), (1318, 688), (1540, 920), (1860, 960)):
                with self.subTest(scale=scale, width=width, height=height):
                    self.app.geometry(f"{width}x{height}+0+0")
                    self.settle()
                    self.assert_filters_visible()
                    table = self.app.overview_tree.master
                    self.assertLessEqual(abs(table.winfo_height() - self.app.overview_summary.winfo_height()), 2)
                    self.assert_inside(self.app.overview_tree, table)

    def test_fullscreen_retains_category_and_sorting_controls(self) -> None:
        self.app.geometry("1318x688+0+0")
        self.settle()
        mode = self.app._table_modes["overview"]
        before = self.app.overview_tree.winfo_height()
        mode.expand()
        self.app._apply_ui_scale(0.9)
        self.settle()
        self.assert_filters_visible()
        self.assertGreater(self.app.overview_tree.winfo_height(), before)
        self.assertFalse(self.app.overview_summary.winfo_viewable())
        mode.restore()
        self.settle()
        self.assert_filters_visible()
        self.assertTrue(self.app.overview_summary.winfo_viewable())

    def test_summary_scrolls_without_moving_the_filters(self) -> None:
        self.app.geometry("1180x608+0+0")
        self.settle()
        summary = self.app.overview_summary
        self.assertIsInstance(summary, ScrollableSummary)
        self.assertTrue(summary.winfo_viewable())
        # Guarantee overflow regardless of font metrics on the runner.
        spacer = ttk.Frame(summary.content, height=summary.canvas.winfo_height() + 150)
        spacer.grid(row=999, column=0)
        self.addCleanup(spacer.destroy)
        self.settle()
        self.assertGreater(summary.content.winfo_reqheight(), summary.canvas.winfo_height())
        self.assertTrue(summary.yscroll.winfo_viewable())
        filters_y = self.app.overview_filters.winfo_rooty()
        summary.canvas.yview_moveto(1)
        self.settle()
        self.assertGreater(summary.canvas.yview()[0], 0)
        self.assertEqual(self.app.overview_filters.winfo_rooty(), filters_y)
        self.assert_filters_visible()


if __name__ == "__main__":
    unittest.main()
