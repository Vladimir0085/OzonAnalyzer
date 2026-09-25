"""Detect Ozon PDF acts whose income is already present in the accrual report."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable

from .excel_reader import REPORT_ACCRUAL, REPORT_ADDITIONAL_INCOME
from .models import AccrualRow, AdditionalIncomeRow, ParsedSource


# Amounts are compared in kopecks: an act matches a row that differs by 0,01 ₽.
AMOUNT_TOLERANCE_KOPECKS = 1


@dataclass(slots=True)
class DoubleCountMatch:
    act_source: ParsedSource
    act: AdditionalIncomeRow
    accrual_rows: tuple[AccrualRow, ...]

    def description(self) -> str:
        row = self.accrual_rows[0]
        accrual_type = row.accrual_type or "Без типа начисления"
        row_date = f"от {row.accrual_date:%d.%m.%Y}" if row.accrual_date else "без даты"
        text = (
            f"Возможен двойной учёт: акт № {self.act.document_number} от "
            f"{self.act.income_date:%d.%m.%Y} на {_money(self.act.amount)} совпадает "
            f"со строкой «{accrual_type}» {row_date} в отчёте по начислениям."
        )
        if len(self.accrual_rows) > 1:
            text += f" Всего совпадающих строк: {len(self.accrual_rows)}."
        return text

    def question(self) -> str:
        return f"{self.description()} Добавить акт всё равно?"

    def quality_message(self) -> str:
        return f"{self.description()} Акт добавлен по решению пользователя."


def find_double_count_matches(sources: Iterable[ParsedSource]) -> list[DoubleCountMatch]:
    """Return PDF acts that repeat report-level income of the same month.

    A PDF act is suspicious when the accrual report contains a positive row
    without an article for the same amount (±0,01 ₽) in the act's calendar month.
    Rows with an article belong to products and are never compared.
    """
    sources = list(sources)
    candidates = [
        row
        for source in sources
        if source.report_type == REPORT_ACCRUAL
        for row in source.accrual_rows
        if not row.article and row.amount > 0
    ]
    matches: list[DoubleCountMatch] = []
    for source in sources:
        if source.report_type != REPORT_ADDITIONAL_INCOME:
            continue
        for act in source.additional_income_rows:
            rows = tuple(
                row
                for row in candidates
                if _same_month(row.accrual_date, act.income_date)
                and _same_amount(row.amount, act.amount)
            )
            if rows:
                matches.append(DoubleCountMatch(source, act, rows))
    return matches


def _same_month(value: date | None, act_date: date) -> bool:
    # A dateless row can only come from the report already matched to the
    # act's month, so it is still a potential duplicate.
    return value is None or (value.year, value.month) == (act_date.year, act_date.month)


def _same_amount(left: float, right: float) -> bool:
    return abs(round(left * 100) - round(right * 100)) <= AMOUNT_TOLERANCE_KOPECKS


def _money(value: float) -> str:
    return f"{value:,.2f} ₽".replace(",", " ")
