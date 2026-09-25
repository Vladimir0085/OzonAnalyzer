from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from ozon_app.calculator import calculate_run
from ozon_app.database import DOUBLE_COUNT_EVENT
from ozon_app.double_count import find_double_count_matches
from ozon_app.excel_reader import REPORT_ACCRUAL, REPORT_ADDITIONAL_INCOME
from ozon_app.models import AccrualRow, AdditionalIncomeRow, ParsedSource, Product
from ozon_app.service import AppService, ImportBatch, split_import_sources
from ozon_app.ui import OZPriceAnalyzerApp


LOSS = "Потеря по вине Ozon в логистике"


def _row(
    amount: float,
    value_date: date | None = date(2026, 6, 12),
    article: str = "",
    accrual_type: str = LOSS,
    row_number: int = 5,
) -> AccrualRow:
    return AccrualRow(
        source_name="june.xlsx",
        sheet_name="Начисления",
        row_number=row_number,
        accrual_id=f"row-{row_number}",
        accrual_date=value_date,
        accrual_type=accrual_type,
        article=article,
        sku="",
        product_name="",
        quantity=0,
        seller_price=0,
        amount=amount,
    )


def _accrual_source(*rows: AccrualRow, path: Path = Path("june.xlsx")) -> ParsedSource:
    dates = [row.accrual_date for row in rows if row.accrual_date is not None]
    return ParsedSource(
        path=path,
        file_hash=f"accrual-{path.name}",
        report_type=REPORT_ACCRUAL,
        sheet_name="Начисления",
        header_row=1,
        accrual_rows=list(rows),
        period_start=min(dates) if dates else None,
        period_end=max(dates) if dates else None,
    )


def _act_source(
    amount: float = 359.10,
    act_date: date = date(2026, 6, 30),
    path: Path = Path("act.pdf"),
) -> ParsedSource:
    return ParsedSource(
        path=path,
        file_hash=f"act-{path.name}",
        report_type=REPORT_ADDITIONAL_INCOME,
        sheet_name="PDF",
        header_row=0,
        additional_income_rows=[
            AdditionalIncomeRow(
                source_name=path.name,
                page_number=1,
                document_number="100001",
                income_date=act_date,
                income_type="Компенсация за утерю",
                amount=amount,
            )
        ],
        period_start=act_date,
        period_end=act_date,
    )


ACT_TEXT = """
Акт о компенсации за утерю # 100001 от 30.06.2026
Ozon: Интернет Решения, ООО
Итого к начислению, руб. 359.10
"""


class _PdfPage:
    def extract_text(self) -> str:
        return ACT_TEXT


class _PdfReader:
    pages = [_PdfPage()]


def _write_accrual_report(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Начисления"
    sheet.append(
        [
            "ID начисления",
            "Дата начисления",
            "Тип начисления",
            "Артикул",
            "SKU",
            "Название товара",
            "Количество",
            "Цена продавца",
            "Сумма итого, руб.",
        ]
    )
    sheet.append(["operation-1", date(2026, 6, 12), LOSS, None, None, None, 0, 0, 359.10])
    workbook.save(path)
    workbook.close()


class FindDoubleCountMatchesTests(unittest.TestCase):
    def test_positive_row_without_article_with_same_amount_matches(self) -> None:
        act = _act_source()
        row = _row(359.10)

        matches = find_double_count_matches([_accrual_source(row), act])

        self.assertEqual(len(matches), 1)
        self.assertIs(matches[0].act_source, act)
        self.assertEqual(matches[0].accrual_rows, (row,))
        self.assertEqual(
            matches[0].question(),
            "Возможен двойной учёт: акт № 100001 от 30.06.2026 на 359.10 ₽ совпадает "
            f"со строкой «{LOSS}» от 12.06.2026 в отчёте по начислениям. "
            "Добавить акт всё равно?",
        )

    def test_amount_within_one_kopeck_matches(self) -> None:
        for amount in (359.09, 359.11):
            with self.subTest(amount=amount):
                matches = find_double_count_matches(
                    [_accrual_source(_row(amount)), _act_source(359.10)]
                )
                self.assertEqual(len(matches), 1)

    def test_different_amount_does_not_match(self) -> None:
        for amount in (359.12, 359.08, 100.00):
            with self.subTest(amount=amount):
                matches = find_double_count_matches(
                    [_accrual_source(_row(amount)), _act_source(359.10)]
                )
                self.assertEqual(matches, [])

    def test_same_amount_in_another_month_does_not_match(self) -> None:
        may_row = _row(359.10, value_date=date(2026, 5, 31))
        june_row = _row(10.00, value_date=date(2026, 6, 1), row_number=6)

        matches = find_double_count_matches(
            [_accrual_source(may_row, june_row), _act_source(act_date=date(2026, 6, 30))]
        )

        self.assertEqual(matches, [])

    def test_row_with_article_does_not_match(self) -> None:
        matches = find_double_count_matches(
            [_accrual_source(_row(359.10, article="A")), _act_source()]
        )

        self.assertEqual(matches, [])

    def test_negative_row_does_not_match(self) -> None:
        matches = find_double_count_matches(
            [_accrual_source(_row(-359.10)), _act_source()]
        )

        self.assertEqual(matches, [])

    def test_several_matching_rows_are_reported_once_per_act(self) -> None:
        first = _row(359.10, row_number=5)
        second = _row(359.10, value_date=date(2026, 6, 20), row_number=9)

        matches = find_double_count_matches([_accrual_source(first, second), _act_source()])

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].accrual_rows, (first, second))
        self.assertIn("от 12.06.2026", matches[0].description())
        self.assertIn("Всего совпадающих строк: 2.", matches[0].description())

    def test_module_does_not_import_tkinter(self) -> None:
        result = subprocess.run(
            [
                sys.executable,
                "-c",
                "import sys, ozon_app.double_count; print('tkinter' in sys.modules)",
            ],
            cwd=Path(__file__).resolve().parents[1],
            capture_output=True,
            text=True,
            check=True,
        )

        self.assertEqual(result.stdout.strip(), "False")


class DoubleCountImportTests(unittest.TestCase):
    def test_declined_act_is_excluded_from_unallocated_income(self) -> None:
        accrual = _accrual_source(_row(359.10))
        act = _act_source()
        session = split_import_sources([act, accrual])[0]

        [match] = session.double_count_matches()
        session.exclude_source(match.act_source)
        calculation = calculate_run(session.sources, {"A": Product("A", "Товар")}, 0.04)

        self.assertEqual(session.sources, [accrual])
        self.assertEqual(session.double_count_matches(), [])
        self.assertAlmostEqual(calculation.unallocated_total, 359.10)

    def test_accrual_report_cannot_be_excluded(self) -> None:
        accrual = _accrual_source(_row(359.10))
        session = split_import_sources([accrual, _act_source()])[0]

        with self.assertRaises(ValueError):
            session.exclude_source(accrual)

    def test_accepted_act_is_counted_and_logged_in_quality_control(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            service = AppService(root / "storage")
            accrual_path = root / "june.xlsx"
            act_path = root / "act.pdf"
            accrual_path.write_bytes(b"synthetic-accrual")
            act_path.write_bytes(b"synthetic-act")
            session = split_import_sources(
                [
                    _accrual_source(_row(359.10), path=accrual_path),
                    _act_source(path=act_path),
                ]
            )[0]
            [match] = session.double_count_matches()

            calculation = service.complete_import(
                session,
                double_count_warnings=[match.quality_message()],
            )

            self.assertAlmostEqual(calculation.unallocated_total, 718.20)
            events = [
                event
                for event in service.db.list_quality_events(calculation.run_id)
                if event["event_type"] == DOUBLE_COUNT_EVENT
            ]
            self.assertEqual(len(events), 1)
            self.assertEqual(events[0]["severity"], "Предупреждение")
            self.assertTrue(events[0]["message"].startswith("Возможен двойной учёт: акт № 100001"))
            self.assertTrue(events[0]["message"].endswith("Акт добавлен по решению пользователя."))
            self.assertEqual(
                service.db.load_calculation(calculation.run_id).double_count_warnings,
                [match.quality_message()],
            )

    def test_real_files_warning_survives_history_recalculation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            service = AppService(root / "storage")
            accrual_path = root / "june.xlsx"
            act_path = root / "act.pdf"
            _write_accrual_report(accrual_path)
            act_path.write_bytes(b"synthetic-pdf")

            with patch("pypdf.PdfReader", return_value=_PdfReader()):
                [session] = service.prepare_import([accrual_path, act_path]).sessions
                [match] = session.double_count_matches()
                calculation = service.complete_import(
                    session,
                    double_count_warnings=[match.quality_message()],
                )
                repair = service.recalculate_history()

            self.assertIn(f"«{LOSS}» от 12.06.2026", match.question())
            self.assertAlmostEqual(calculation.unallocated_total, 718.20)
            new_id = repair.old_to_new[calculation.run_id]
            self.assertEqual(
                service.db.load_calculation(new_id).double_count_warnings,
                [match.quality_message()],
            )


class _ServiceStub:
    def __init__(self) -> None:
        self.saved: dict[str, object] = {}

    def replacement_run_ids(self, _session) -> list[int]:
        return []

    def complete_import_batch(self, sessions, **kwargs):
        self.saved = {"sources": [list(session.sources) for session in sessions], **kwargs}
        return [calculate_run(session.sources, {}, 0.04) for session in sessions]


class _DbStub:
    def get_setting(self, _key: str, default: str) -> str:
        return "0"

    def product_map(self, active_only: bool = True) -> dict:
        return {}


class _StatusStub:
    def set(self, _value: str) -> None:
        pass


class _ImportAppStub:
    run_number_by_id: dict[int, int] = {}

    def __init__(self) -> None:
        self.service = _ServiceStub()
        self.db = _DbStub()
        self.status_var = _StatusStub()
        self.import_in_progress = True
        self.current_run_id = None

    def refresh_all(self) -> None:
        pass

    def configure(self, **_kwargs) -> None:
        pass

    def _current_run_status(self) -> str:
        return ""


class DoubleCountDialogTests(unittest.TestCase):
    def _run_import(self, answer: bool):
        accrual = _accrual_source(_row(359.10))
        act = _act_source()
        batch = ImportBatch(sessions=split_import_sources([accrual, act]), unknown_products=[])
        app = _ImportAppStub()
        with (
            patch("ozon_app.ui.messagebox.askyesno", return_value=answer) as ask,
            patch("ozon_app.ui.messagebox.showinfo") as info,
            patch("ozon_app.ui.messagebox.showerror") as error,
        ):
            OZPriceAnalyzerApp._complete_import_ui(app, batch, None)
        error.assert_not_called()
        return app.service.saved, ask, info, accrual, act

    def test_dialog_defaults_to_no_and_declined_act_is_not_saved(self) -> None:
        saved, ask, info, accrual, _act = self._run_import(answer=False)

        ask.assert_called_once()
        self.assertEqual(ask.call_args.kwargs["default"], "no")
        self.assertTrue(
            ask.call_args.args[1].startswith(
                "Возможен двойной учёт: акт № 100001 от 30.06.2026 на 359.10 ₽"
            )
        )
        self.assertIn("Добавить акт всё равно?", ask.call_args.args[1])
        self.assertEqual(saved["sources"], [[accrual]])
        self.assertEqual(saved["double_count_warnings_by_session"], [[]])
        self.assertIn("Не добавлены PDF-акты", info.call_args.args[1])
        self.assertIn("№ 100001", info.call_args.args[1])

    def test_confirmed_act_is_saved_with_quality_warning(self) -> None:
        saved, _ask, info, accrual, act = self._run_import(answer=True)

        self.assertEqual(saved["sources"], [[accrual, act]])
        [[warning]] = saved["double_count_warnings_by_session"]
        self.assertTrue(warning.startswith("Возможен двойной учёт: акт № 100001"))
        self.assertTrue(warning.endswith("Акт добавлен по решению пользователя."))
        self.assertNotIn("Не добавлены PDF-акты", info.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
