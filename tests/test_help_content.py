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
            "Итог после налога = Чистая прибыль товаров + Нераспределённые",
            text,
        )
        self.assertIn(
            "Доходность итога = Итог с нераспределёнными после налога ÷ С/с проданного × 100%",
            text,
        )
        self.assertIn(
            "Чистая прибыль, % от выручки = Итог с нераспределёнными",
            text,
        )
        self.assertIn("Одинаковый активный набор используют вкладки", text)
        self.assertIn("изменять и сохранять плановые цены можно при выборе одного отчёта", text)

    def test_base_ui_exposes_help_tab_builder(self) -> None:
        self.assertTrue(callable(getattr(OZPriceAnalyzerApp, "_build_help_tab", None)))


if __name__ == "__main__":
    unittest.main()
