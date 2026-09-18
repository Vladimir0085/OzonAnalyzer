from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import patch

from ozon_app.catalog_category_filters import CatalogAndCategoryOZPriceAnalyzerApp
from ozon_app.ui import CATEGORY_ALL, SORT_ASCENDING, SORT_NONE, OZPriceAnalyzerApp


class _Variable:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class _Tree:
    def __init__(self) -> None:
        self.rows: list[tuple[str | None, tuple[object, ...], tuple[str, ...]]] = []
        self.selected: str | None = None
        self.focused: str | None = None

    def get_children(self) -> tuple[str, ...]:
        return tuple(str(index) for index in range(len(self.rows)))

    def delete(self, *_items: str) -> None:
        self.rows.clear()

    def insert(
        self,
        _parent: str,
        _position: str,
        *,
        iid: str | None = None,
        values: tuple[object, ...],
        tags: tuple[str, ...] = (),
    ) -> None:
        self.rows.append((iid, values, tags))

    def selection_set(self, iid: str) -> None:
        self.selected = iid

    def focus(self, iid: str) -> None:
        self.focused = iid


class _Combo(dict):
    pass


class _SourceDatabase:
    def __init__(self) -> None:
        self.calls: list[int] = []

    def list_source_files(self, run_id: int) -> list[dict[str, object]]:
        self.calls.append(run_id)
        return [
            {
                "id": run_id * 10,
                "original_name": f"report-{run_id}.xlsx",
                "report_type": "ACCRUAL",
                "row_count": run_id,
                "total_amount": 100.0 * run_id,
                "period_start": "2026-09-01",
                "period_end": "2026-09-30",
                "file_hash": f"hash-{run_id}",
            }
        ]


class _ScenarioDatabase:
    def __init__(self) -> None:
        self.planned_price_calls: list[int] = []

    def planned_prices(self, run_id: int) -> dict[str, float]:
        self.planned_price_calls.append(run_id)
        return {}


class _Calculation:
    def __init__(self, revenue: float) -> None:
        self.products: list[object] = []
        self.tax_rate = 0.06
        self.unallocated_total = revenue
        self.unallocated = {"Операция": (1, revenue)}
        self.taxable_unallocated_income = 0.0
        self.unallocated_income_tax = 0.0
        self._revenue = revenue

    def totals(self) -> dict[str, float]:
        return {"revenue": self._revenue, "units": 0.0}


class ActiveReportSelectionTests(unittest.TestCase):
    def test_top_selector_replaces_an_earlier_multi_selection(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.run_var = _Variable("second")
        app.run_display_to_id = {"second": 2}
        app.overview_selection_explicit = True
        app.overview_run_ids = {1, 2}
        selected: list[int] = []
        app.select_run = selected.append

        app._on_run_selected()

        self.assertFalse(app.overview_selection_explicit)
        self.assertEqual(app.overview_run_ids, {2})
        self.assertEqual(selected, [2])

    def test_multi_selector_refreshes_every_selection_dependent_tab(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.db = SimpleNamespace(list_runs=lambda: [SimpleNamespace(id=1), SimpleNamespace(id=2)])
        app.overview_run_ids = {1}
        app.current_run_id = 1
        app.overview_calculation = SimpleNamespace(period_start=None, period_end=None)
        app.status_var = _Variable()
        app.wait_window = lambda _dialog: None
        refreshed: list[bool] = []
        app._refresh_active_report_views = lambda: refreshed.append(True)
        dialog = SimpleNamespace(confirmed=True, selected_run_ids={1, 2})

        with patch("ozon_app.ui.OverviewReportSelectionDialog", return_value=dialog):
            app.choose_overview_reports()

        self.assertTrue(app.overview_selection_explicit)
        self.assertEqual(app.overview_run_ids, {1, 2})
        self.assertEqual(refreshed, [True])

    def test_sources_are_loaded_for_all_active_reports(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.current_run_id = 2
        app.overview_run_ids = {1, 2}
        app.overview_runs = [SimpleNamespace(id=1), SimpleNamespace(id=2)]
        app.db = _SourceDatabase()
        app.source_tree = _Tree()
        app.source_by_iid = {}
        app._on_source_selected = lambda: None
        app.clear_xlsx_preview = lambda: None

        app._populate_sources()

        self.assertEqual(app.db.calls, [1, 2])
        self.assertEqual([row[0] for row in app.source_tree.rows], ["1:10", "2:20"])
        self.assertEqual(app.source_tree.selected, "1:10")

    def test_breakdown_uses_the_active_overview_calculation(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.overview_calculation = _Calculation(30.0)
        app.current_calculation = _Calculation(999.0)
        app.breakdown_tree = _Tree()
        app._configure_value_tags = lambda _tree: None

        app._populate_breakdown()

        self.assertEqual(app.breakdown_tree.rows[0][1][2], "30.00 ₽")
        self.assertEqual(app.breakdown_tree.rows[-1][1][2], "30.00 ₽")

    def test_scenario_uses_aggregate_values_without_leaking_saved_prices(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.current_run_id = 2
        app.overview_run_ids = {1, 2}
        app.overview_runs = [SimpleNamespace(id=1), SimpleNamespace(id=2)]
        app.overview_calculation = _Calculation(300.0)
        app.current_calculation = _Calculation(999.0)
        app.db = _ScenarioDatabase()
        app.scenario_tree = _Tree()
        app.scenario_rows = {}
        app.scenario_kpi_vars = {
            "current_revenue": _Variable(),
            "planned_revenue": _Variable(),
            "planned_net": _Variable(),
            "planned_margin": _Variable(),
        }
        app.scenario_category_combo = _Combo()
        app.scenario_category_var = _Variable(CATEGORY_ALL)
        app.scenario_article_var = _Variable("")
        app.scenario_sort_var = _Variable(SORT_NONE)
        app.scenario_sort_direction_var = _Variable(SORT_ASCENDING)
        app.scenario_count_var = _Variable()
        app._configure_value_tags = lambda _tree: None

        app._populate_scenario()

        self.assertEqual(app.scenario_kpi_vars["current_revenue"].value, "300.00 ₽")
        self.assertEqual(app.db.planned_price_calls, [])
        self.assertIn("2 отчета", app.scenario_count_var.value)

    def test_final_app_category_scenario_uses_the_same_active_scope(self) -> None:
        app = object.__new__(CatalogAndCategoryOZPriceAnalyzerApp)
        app.current_run_id = 2
        app.overview_run_ids = {1, 2}
        app.overview_runs = [SimpleNamespace(id=1), SimpleNamespace(id=2)]
        app.overview_calculation = _Calculation(300.0)
        app.current_calculation = _Calculation(999.0)
        app.db = _ScenarioDatabase()
        app.scenario_tree = _Tree()
        app.scenario_rows = {}
        app.scenario_kpi_vars = {
            "current_revenue": _Variable(),
            "planned_revenue": _Variable(),
            "planned_net": _Variable(),
            "planned_margin": _Variable(),
        }
        app.scenario_article_var = _Variable("")
        app.scenario_sort_var = _Variable(SORT_NONE)
        app.scenario_sort_direction_var = _Variable(SORT_ASCENDING)
        app.scenario_count_var = _Variable()
        app._sync_filter_categories = lambda _key, _values: None
        app._filter_scope = lambda _key: (set(), False)
        app._configure_value_tags = lambda _tree: None

        app._populate_scenario()

        self.assertEqual(app.scenario_kpi_vars["current_revenue"].value, "300.00 ₽")
        self.assertEqual(app.db.planned_price_calls, [])
        self.assertIn("2 отчета", app.scenario_count_var.value)

    def test_multi_report_scenario_cannot_save_a_price_to_the_current_run(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.current_run_id = 2
        app.overview_run_ids = {1, 2}
        app.overview_runs = [SimpleNamespace(id=1), SimpleNamespace(id=2)]
        app.scenario_tree = SimpleNamespace(selection=lambda: ("A-1",))

        with patch("ozon_app.ui.messagebox.showinfo") as showinfo:
            app.apply_planned_price()

        showinfo.assert_called_once()
        self.assertIn("одного отчета", showinfo.call_args.args[1].casefold())


if __name__ == "__main__":
    unittest.main()
