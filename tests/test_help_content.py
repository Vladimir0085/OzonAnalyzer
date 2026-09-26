from __future__ import annotations

import unittest

from ozon_app.help_content import (
    OVERVIEW_HELP_CONTENT,
    REPORTS_HELP_CONTENT,
    help_plain_text,
)
from ozon_app.ui import OZPriceAnalyzerApp


class HelpContentTests(unittest.TestCase):
    def test_reports_help_identifies_required_and_conditional_sources(self) -> None:
        text = help_plain_text(REPORTS_HELP_CONTENT)

        self.assertIn("Отчёт по начислениям — основной и обязательный", text)
        self.assertIn("Детализация по выкупленным товарам", text)
        self.assertIn("RealizationReportCIS", text)
        self.assertIn("BuyoutDetails", text)
        self.assertIn("Акт о премии", text)
        self.assertIn("CompensationReport.xlsx", text)
        self.assertIn("учтён дважды", text)
        self.assertIn("Защита от двойного учёта", text)
        self.assertIn("положительная строка без артикула на ту же сумму (±0,01 ₽)", text)
        self.assertIn("По умолчанию акт не добавляется", text)
        self.assertIn("«Контроле качества»", text)

    def test_overview_help_documents_key_formulas(self) -> None:
        text = help_plain_text(OVERVIEW_HELP_CONTENT)

        self.assertIn(
            "Выручка = Выручка без баллов + Программы партнёров + Баллы",
            text,
        )
        self.assertIn(
            "Чистая прибыль = Финрезультат Ozon − С/с проданного − Налог товара",
            text,
        )
        self.assertIn(
            "Чистая прибыль от деятельности = Чистая прибыль товаров + Нераспределённые − "
            "Налог с положительных нераспределённых доходов",
            text,
        )
        self.assertIn(
            "Доходность деятельности = Чистая прибыль от деятельности ÷ С/с проданного × 100%",
            text,
        )
        self.assertIn(
            "Чистая прибыль, % от выручки = Чистая прибыль от деятельности ÷ Выручка",
            text,
        )
        self.assertIn(
            "Логистика, % = (Обработка отправления + Доставка до ПВЗ + Логистика + "
            "Обратная логистика + Возвраты/отмены)",
            text,
        )
        self.assertIn("Одинаковый активный набор используют вкладки", text)
        self.assertIn("изменять и сохранять плановые цены можно при выборе одного отчёта", text)

    def test_help_documents_tax_rates_by_period(self) -> None:
        text = help_plain_text(OVERVIEW_HELP_CONTENT)

        self.assertIn("Налоговые ставки по периодам", text)
        self.assertIn(
            "Ставка на дату = ставка строки с наибольшей датой «Действует с», не позже этой даты",
            text,
        )
        self.assertIn("Выручка, Возврат выручки, Программы партнёров — по дате начисления", text)
        self.assertIn("по дате окончания периода отчёта о выкупах", text)
        self.assertIn("PDF-акты — по дате акта", text)
        self.assertIn("Баллы за скидки в налоговую базу не входят", text)
        self.assertIn("Объединённый обзор складывает сохранённые налоги отчётов", text)
        self.assertIn("ставке на дату окончания периода отчёта", text)
        self.assertIn("ставка самого позднего из них", text)
        self.assertIn("При отмене не меняется ничего", text)
        self.assertIn("«Дата» в группе «Отчёт о выкупленных товарах»", help_plain_text(REPORTS_HELP_CONTENT))

    def test_base_ui_exposes_help_tab_builder(self) -> None:
        self.assertTrue(callable(getattr(OZPriceAnalyzerApp, "_build_help_tab", None)))


if __name__ == "__main__":
    unittest.main()
