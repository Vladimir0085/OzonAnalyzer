"""Display-only labels shared by the interface and the XLSX export.

These helpers never change a calculated number: they only choose the text
shown in place of a profitability percentage whose denominator is missing.
"""

from __future__ import annotations


NO_SALES_TEXT = "Нет продаж"
NO_COST_TEXT = "Нет себестоимости"


def profitability_label(units: float, cost_sold: float) -> str | None:
    """Label shown instead of profitability (the WB 0.2.15 rule).

    No sales (units ≤ 0) → «Нет продаж»; otherwise no cost of goods sold
    (≤ 0) → «Нет себестоимости»; otherwise ``None`` and the percentage is shown.
    """
    if units <= 0:
        return NO_SALES_TEXT
    if cost_sold <= 0:
        return NO_COST_TEXT
    return None
