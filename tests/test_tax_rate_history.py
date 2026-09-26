from __future__ import annotations

import sqlite3
import tempfile
import unittest
from contextlib import closing
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from ozon_app.backup import create_backup, restore_backup
from ozon_app.database import Database
from ozon_app.models import ParsedSource, Product, ProductResult, RunCalculation
from ozon_app.service import AppService
from ozon_app.tax_rates import TaxRatePeriod, TaxRateSchedule
from ozon_app.ui import (
    OZPriceAnalyzerApp,
    parse_tax_rate_input,
    scenario_tax_note,
    tax_rate_table_values,
    tax_recalculation_rows,
)


MID_AUGUST = TaxRateSchedule(
    [TaxRatePeriod(None, 0.04), TaxRatePeriod(date(2026, 8, 15), 0.06)]
)
HEADERS = [
    "ID начисления",
    "Дата начисления",
    "Группа услуг",
    "Тип начисления",
    "Артикул",
    "SKU",
    "Название товара",
    "Количество",
    "Цена продавца",
    "Сумма итого, руб.",
]


def _write_accrual_report(path: Path, rows: list[tuple[object, ...]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Начисления"
    sheet.append(HEADERS)
    for row in rows:
        sheet.append(list(row))
    workbook.save(path)
    workbook.close()


def _sale(operation: str, day: date, amount: float) -> tuple[object, ...]:
    return (operation, day, "Продажи", "Выручка", "A", 1001, "Тестовый товар", 1, amount, amount)


def _unallocated(operation: str, day: date, accrual_type: str, amount: float) -> tuple[object, ...]:
    return (operation, day, "", accrual_type, "", "", "", 0, 0, amount)


class _Variable:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class TaxRateRecalculationTests(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.service = AppService(self.root / "data")
        self.service.db.save_product(Product("A", "Тестовый товар", 60, 40), source="Тест")
        july = self.root / "Отчет по начислениям_01.07.2026-31.07.2026.xlsx"
        august = self.root / "Отчет по начислениям_01.08.2026-31.08.2026.xlsx"
        _write_accrual_report(
            july,
            [
                _sale("J-1", date(2026, 7, 1), 1_000),
                _unallocated("J-2", date(2026, 7, 31), "Подписка Premium", -10),
            ],
        )
        _write_accrual_report(
            august,
            [
                _sale("A-1", date(2026, 8, 1), 1_000),
                _sale("A-2", date(2026, 8, 20), 2_000),
                _unallocated("A-3", date(2026, 8, 25), "Премия Ozon", 100),
                _unallocated("A-4", date(2026, 8, 31), "Подписка Premium", -10),
            ],
        )
        batch = self.service.prepare_import([july, august])
        self.july, self.august = self.service.complete_import_batch(batch.sessions)
        self.service.db.rename_run(self.august.run_id, "Август — проверено")
        self.service.db.save_planned_price(self.august.run_id, "A", 1_700.0)

    def _run_ids(self) -> list[int]:
        return [run.id for run in self.service.db.list_runs()]

    def test_initial_schedule_comes_from_settings_and_import_uses_it(self) -> None:
        self.assertEqual(self.service.db.tax_rate_schedule(), TaxRateSchedule.single(0.04))
        self.assertAlmostEqual(self.august.totals()["tax"], 124)

    def test_preview_lists_only_affected_reports_and_changes_nothing(self) -> None:
        before_ids = self._run_ids()

        plan = self.service.preview_tax_rate_change(MID_AUGUST)

        [item] = plan.changed_items
        self.assertEqual(item.run_id, self.august.run_id)
        self.assertEqual(item.report_name, "Август — проверено")
        self.assertAlmostEqual(item.old_tax, 40 + 80 + 4)
        self.assertAlmostEqual(item.new_tax, 40 + 120 + 6)
        self.assertAlmostEqual(item.product_net_profit_delta, -40)
        self.assertAlmostEqual(item.activity_net_profit_delta, -42)
        self.assertEqual(plan.unchanged_tax_count, 0)
        self.assertEqual(
            tax_recalculation_rows(plan, {self.august.run_id: 2}),
            [
                (
                    "№2 · Август — проверено",
                    "01.08.2026–31.08.2026",
                    "124.00 ₽",
                    "166.00 ₽",
                    "-40.00 ₽",
                    "-42.00 ₽",
                )
            ],
        )
        # Без подтверждения ничего не меняется.
        self.assertEqual(self.service.db.tax_rate_schedule(), TaxRateSchedule.single(0.04))
        self.assertEqual(self._run_ids(), before_ids)
        self.assertAlmostEqual(
            self.service.db.load_calculation(self.august.run_id).totals()["tax"], 124
        )
        self.assertEqual(list((self.root / "data" / "backups").iterdir()), [])

    def test_apply_recalculates_with_backup_and_preserves_history_inputs(self) -> None:
        old_created_at = {run.id: run.created_at for run in self.service.db.list_runs()}
        plan = self.service.preview_tax_rate_change(MID_AUGUST)
        backup = self.root / "data" / "backups" / "before.ozbackup"

        old_to_new = self.service.apply_tax_rate_change(plan, backup)

        self.assertTrue(backup.is_file())
        self.assertEqual(set(old_to_new), {self.august.run_id})
        new_id = old_to_new[self.august.run_id]
        self.assertEqual(self.service.db.tax_rate_schedule(), MID_AUGUST)
        runs = {run.id: run for run in self.service.db.list_runs()}
        self.assertEqual(set(runs), {self.july.run_id, new_id})
        self.assertEqual(runs[new_id].report_name, "Август — проверено")
        self.assertEqual(runs[new_id].created_at, old_created_at[self.august.run_id])
        self.assertEqual(self.service.db.planned_prices(new_id), {"A": 1_700.0})

        august = self.service.db.load_calculation(new_id)
        self.assertAlmostEqual(august.products[0].tax(august.tax_rate), 160)
        self.assertAlmostEqual(august.unallocated_income_tax, 6)
        self.assertEqual(august.products[0].material_cost, 60)
        self.assertEqual(august.tax_rate, 0.06)
        self.assertEqual(self.service.db.run_tax_schedule(new_id), MID_AUGUST)
        july = self.service.db.load_calculation(self.july.run_id)
        self.assertAlmostEqual(july.totals()["tax"], 40)

        # Повторное применение того же справочника ничего не находит,
        # а «Пересчитать историю» использует сохраненные ставки отчетов.
        self.assertEqual(self.service.preview_tax_rate_change(MID_AUGUST).items, [])
        repair = self.service.recalculate_history()
        recalculated = self.service.db.load_calculation(repair.old_to_new[new_id])
        self.assertAlmostEqual(recalculated.totals()["tax"], 166)

    def test_schedule_change_without_affected_reports_saves_without_backup(self) -> None:
        future = TaxRateSchedule(
            [TaxRatePeriod(None, 0.04), TaxRatePeriod(date(2030, 1, 1), 0.06)]
        )
        before_ids = self._run_ids()
        plan = self.service.preview_tax_rate_change(future)
        backup = self.root / "data" / "backups" / "unused.ozbackup"

        self.assertEqual(plan.items, [])
        self.assertEqual(self.service.apply_tax_rate_change(plan, backup), {})
        self.assertFalse(backup.exists())
        self.assertEqual(self.service.db.tax_rate_schedule(), future)
        self.assertEqual(self._run_ids(), before_ids)

    def test_stale_preview_and_missing_sources_leave_everything_unchanged(self) -> None:
        plan = self.service.preview_tax_rate_change(MID_AUGUST)
        self.service.db.replace_tax_rates(TaxRateSchedule.single(0.05))
        with self.assertRaisesRegex(ValueError, "после предпросмотра"):
            self.service.apply_tax_rate_change(plan)
        self.assertEqual(self.service.db.tax_rate_schedule(), TaxRateSchedule.single(0.05))

        self.service.db.replace_tax_rates(TaxRateSchedule.single(0.04))
        for record in self.service.db.list_source_files(self.august.run_id):
            Path(str(record["stored_path"])).unlink()
        with self.assertRaisesRegex(ValueError, "Не найден сохраненный исходный файл"):
            self.service.preview_tax_rate_change(MID_AUGUST)
        self.assertEqual(self.service.db.tax_rate_schedule(), TaxRateSchedule.single(0.04))

    def test_settings_flow_cancel_keeps_data_and_confirm_applies(self) -> None:
        app = object.__new__(OZPriceAnalyzerApp)
        app.service = self.service
        app.db = self.service.db
        app.status_var = _Variable()
        app.configure = lambda **_options: None
        app.update_idletasks = lambda: None
        app.wait_window = lambda _dialog: None
        app.run_number_by_id = {}
        app._current_run_status = lambda: "Готово"
        app._refresh_tax_rates = lambda: None
        followed: list[dict[int, int]] = []
        app._follow_replaced_runs = followed.append
        decisions: list[bool] = []

        class _PreviewDialog:
            def __init__(self, _parent, plan, _numbers) -> None:
                self.plan = plan
                self.confirmed = decisions.pop(0)

        periods = list(MID_AUGUST.periods)
        with patch("ozon_app.ui.TaxRecalculationPreviewDialog", _PreviewDialog), patch(
            "ozon_app.ui.messagebox"
        ) as messagebox:
            decisions.append(False)
            app._change_tax_rates(periods)
            self.assertEqual(self.service.db.tax_rate_schedule(), TaxRateSchedule.single(0.04))
            self.assertIn(self.august.run_id, self._run_ids())
            self.assertIn("отменено", app.status_var.value)

            decisions.append(True)
            app._change_tax_rates(periods)
            messagebox.showerror.assert_not_called()

        self.assertEqual(self.service.db.tax_rate_schedule(), MID_AUGUST)
        [mapping] = followed
        self.assertEqual(set(mapping), {self.august.run_id})
        backups = list((self.root / "data" / "backups").glob("Автокопия_перед_изменением_ставок_*.ozbackup"))
        self.assertEqual(len(backups), 1)


class LegacyTaxMigrationTests(unittest.TestCase):
    def test_saved_reports_keep_their_results_after_migration(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "legacy.sqlite3"
            database = Database(path)
            source_path = root / "stored.xlsx"
            source_path.write_bytes(b"source")
            run_ids = []
            for month, rate in ((7, 0.04), (8, 0.06)):
                source = ParsedSource(
                    path=source_path,
                    file_hash=f"hash-{month}",
                    report_type="REALIZATION",
                    sheet_name="Детализация",
                    header_row=10,
                    period_start=date(2026, month - 1, 16),
                    period_end=date(2026, month - 1, 30),
                )
                run_ids.append(
                    database.save_run(
                        RunCalculation(
                            run_id=None,
                            period_start=date(2026, month, 1),
                            period_end=date(2026, month, 31),
                            tax_rate=rate,
                            products=[
                                ProductResult(
                                    "A",
                                    "Товар",
                                    60.1,
                                    40.3,
                                    units=3,
                                    revenue_no_points=1_234.57,
                                    partner_programs=98.76,
                                    points=55.5,
                                    commission=-201.01,
                                    financial_result=987.65,
                                ),
                                ProductResult(
                                    "B",
                                    "Товар 2",
                                    10,
                                    5,
                                    units=1,
                                    revenue_no_points=333.33,
                                    financial_result=111.11,
                                ),
                            ],
                            unallocated_total=140.78,
                            unallocated={
                                "Брак по вине Ozon на складе": (1, 120.78),
                                "Премия Ozon": (1, 40.0),
                                "Подписка": (1, -20.0),
                            },
                            accrual_stats={},
                            source_files=[source],
                        ),
                        {str(source_path): source_path},
                    )
                )
            database.set_setting("tax_rate", "0.06")

            def snapshot(db: Database) -> list[object]:
                result: list[object] = [
                    (run.units, run.revenue, run.net_profit, run.profitability, run.net_margin)
                    for run in db.list_runs()
                ]
                for run_id in run_ids:
                    calculation = db.load_calculation(run_id)
                    result.append(sorted(calculation.totals().items()))
                    result.append(calculation.report_net_profit)
                    result.append(
                        [
                            (item.tax(calculation.tax_rate), item.net_profit(calculation.tax_rate))
                            for item in calculation.products
                        ]
                    )
                return result

            before = snapshot(database)
            with closing(sqlite3.connect(path)) as connection:
                for statement in (
                    "ALTER TABLE product_results DROP COLUMN tax",
                    "ALTER TABLE runs DROP COLUMN unallocated_income_tax",
                    "ALTER TABLE runs DROP COLUMN tax_period_start",
                    "ALTER TABLE runs DROP COLUMN tax_period_end",
                    "DROP TABLE tax_rates",
                    "DROP TABLE run_tax_rates",
                ):
                    connection.execute(statement)
                connection.commit()

            migrated = Database(path)

            self.assertEqual(snapshot(migrated), before)
            self.assertEqual(migrated.tax_rate_schedule(), TaxRateSchedule.single(0.06))
            self.assertEqual(migrated.run_tax_schedule(run_ids[0]), TaxRateSchedule.single(0.04))
            self.assertEqual(migrated.run_tax_schedule(run_ids[1]), TaxRateSchedule.single(0.06))
            july = migrated.load_calculation(run_ids[0])
            self.assertEqual(
                (july.tax_period_start, july.tax_period_end),
                (date(2026, 6, 16), date(2026, 7, 31)),
            )
            self.assertEqual(july.unallocated_income_tax, 160.78 * 0.04)
            self.assertEqual(
                july.products[0].tax(july.tax_rate),
                (1_234.57 + 98.76) * 0.04,
            )


class TaxRateBackupTests(unittest.TestCase):
    def test_backup_restores_schedule_and_old_backup_gets_one_from_settings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            service = AppService(root / "data")
            service.db.replace_tax_rates(MID_AUGUST)
            archive = create_backup(service.paths["root"], root / "copy.ozbackup").path
            service.db.replace_tax_rates(TaxRateSchedule.single(0.1))

            restore_backup(service.paths["root"], archive)

            self.assertEqual(Database(service.paths["database"]).tax_rate_schedule(), MID_AUGUST)

            legacy = AppService(root / "legacy")
            legacy.db.set_setting("tax_rate", "0.05")
            with closing(sqlite3.connect(legacy.paths["database"])) as connection:
                connection.execute("DROP TABLE tax_rates")
                connection.execute("DROP TABLE run_tax_rates")
                connection.commit()
            legacy_archive = create_backup(legacy.paths["root"], root / "legacy.ozbackup").path

            restore_backup(service.paths["root"], legacy_archive)

            self.assertEqual(
                Database(service.paths["database"]).tax_rate_schedule(),
                TaxRateSchedule.single(0.05),
            )


class TaxRateInputTests(unittest.TestCase):
    def test_dialog_input_validation_and_labels(self) -> None:
        self.assertEqual(
            parse_tax_rate_input("01.08.2026", "6,5", opening=False),
            TaxRatePeriod(date(2026, 8, 1), 0.065),
        )
        self.assertEqual(
            parse_tax_rate_input("с начала", "4", opening=True),
            TaxRatePeriod(None, 0.04),
        )
        for date_text, rate_text in (("2026-08-01", "6"), ("", "6"), ("01.08.2026", "101"), ("01.08.2026", "x")):
            with self.subTest(date_text=date_text, rate_text=rate_text), self.assertRaises(ValueError):
                parse_tax_rate_input(date_text, rate_text, opening=False)
        self.assertEqual(tax_rate_table_values(TaxRatePeriod(None, 0.04)), ("с начала", "4"))
        self.assertEqual(
            tax_rate_table_values(TaxRatePeriod(date(2026, 8, 15), 0.065)),
            ("15.08.2026", "6,5"),
        )
        calculation = RunCalculation(
            run_id=None,
            period_start=date(2026, 8, 1),
            period_end=date(2026, 8, 31),
            tax_rate=0.06,
            products=[],
            unallocated_total=0,
            unallocated={},
            accrual_stats={},
        )
        self.assertEqual(
            scenario_tax_note(calculation),
            "Налог в сценарии: 6 % (ставка на 31.08.2026)",
        )


if __name__ == "__main__":
    unittest.main()
