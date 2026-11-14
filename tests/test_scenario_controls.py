from __future__ import annotations

import unittest
from unittest.mock import patch

from ozon_app.ui import OZPriceAnalyzerApp


class _Variable:
    def __init__(self, value: str):
        self.value = value

    def get(self) -> str:
        return self.value


class _ScenarioTree:
    def __init__(self, selection: tuple[str, ...]):
        self.current_selection = selection
        self.selected: str | None = None
        self.focused: str | None = None
        self.visible: str | None = None

    def selection(self) -> tuple[str, ...]:
        return self.current_selection

    def selection_set(self, article: str) -> None:
        self.selected = article
        self.current_selection = (article,)

    def focus(self, article: str) -> None:
        self.focused = article

    def see(self, article: str) -> None:
        self.visible = article

    def exists(self, article: str) -> bool:
        return article in {"A-1", "A-2"}


class _Database:
    def __init__(self) -> None:
        self.saved: list[tuple[int, str, float]] = []

    def save_planned_price(self, run_id: int, article: str, price: float) -> None:
        self.saved.append((run_id, article, price))


class _ScenarioRow:
    def __init__(self, current_price: float | None):
        self.current_price = current_price


class ScenarioControlTests(unittest.TestCase):
    def test_absolute_price_is_saved_for_one_position_and_selection_is_restored(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.scenario_tree = _ScenarioTree(("A-2",))
        app.db = _Database()
        app._active_single_run_id = lambda notify=False: 12

        def populate() -> None:
            app.scenario_tree.current_selection = ()

        app._populate_scenario = populate

        self.assertTrue(app._save_planned_price("A-2", "2 500,50"))
        self.assertEqual(app.db.saved, [(12, "A-2", 2500.5)])
        self.assertEqual(app.scenario_tree.selected, "A-2")
        self.assertEqual(app.scenario_tree.focused, "A-2")
        self.assertEqual(app.scenario_tree.visible, "A-2")

    def test_negative_absolute_price_is_rejected(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.scenario_tree = _ScenarioTree(("A-2",))
        app.db = _Database()
        app._active_single_run_id = lambda notify=False: 12
        app._populate_scenario = lambda: None

        with patch("ozon_app.ui.messagebox.showerror") as showerror:
            self.assertFalse(app._save_planned_price("A-2", "-1"))

        self.assertEqual(app.db.saved, [])
        showerror.assert_called_once()

    def test_percent_is_applied_only_to_highlighted_row_and_selection_is_restored(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.current_run_id = 12
        app.scenario_tree = _ScenarioTree(("A-2",))
        app.scenario_rows = {
            "A-1": _ScenarioRow(100),
            "A-2": _ScenarioRow(200),
        }
        app.batch_percent_var = _Variable("10")
        app.db = _Database()

        def populate() -> None:
            app.scenario_tree.current_selection = ()

        app._populate_scenario = populate
        app.apply_selected_percent()

        self.assertEqual(app.db.saved[0][:2], (12, "A-2"))
        self.assertAlmostEqual(app.db.saved[0][2], 220.0)
        self.assertEqual(app.scenario_tree.selected, "A-2")
        self.assertEqual(app.scenario_tree.focused, "A-2")
        self.assertEqual(app.scenario_tree.visible, "A-2")

    def test_selected_percent_requires_a_highlighted_row(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.current_run_id = 12
        app.scenario_tree = _ScenarioTree(())
        app.scenario_rows = {}
        app.batch_percent_var = _Variable("10")
        app.db = _Database()

        with patch("ozon_app.ui.messagebox.showinfo") as showinfo:
            app.apply_selected_percent()

        self.assertEqual(app.db.saved, [])
        showinfo.assert_called_once()
        self.assertIn("выберите товар", showinfo.call_args.args[1].casefold())


if __name__ == "__main__":
    unittest.main()
