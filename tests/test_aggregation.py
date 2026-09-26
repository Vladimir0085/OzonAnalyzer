from __future__ import annotations

import unittest
from datetime import date

from ozon_app.aggregation import aggregate_calculations
from ozon_app.models import ProductResult, RunCalculation


def _calculation(
    run_id: int,
    value_date: date,
    tax_rate: float,
    product: ProductResult,
    unallocated: float,
) -> RunCalculation:
    return RunCalculation(
        run_id=run_id,
        period_start=value_date.replace(day=1),
        period_end=value_date,
        tax_rate=tax_rate,
        products=[product],
        unallocated_total=unallocated,
        unallocated={"Подписка": (1, unallocated)},
        accrual_stats={"Подписка": (0, 1)},
        realization_revenue=50 * run_id,
        realization_units=run_id,
    )


class AggregationTests(unittest.TestCase):
    def test_combines_months_with_exact_historical_costs_and_taxes(self) -> None:
        april = _calculation(
            1,
            date(2026, 4, 30),
            0.04,
            ProductResult(
                "A",
                "Товар",
                60,
                40,
                category="Горшки",
                units=2,
                revenue_no_points=500,
                financial_result=300,
            ),
            -10,
        )
        may = _calculation(
            2,
            date(2026, 5, 31),
            0.06,
            ProductResult(
                "A",
                "Товар новое имя",
                80,
                20,
                category="Новая категория",
                units=3,
                revenue_no_points=900,
                financial_result=600,
            ),
            -15,
        )

        combined = aggregate_calculations([may, april])
        row = combined.products[0]

        self.assertEqual(combined.period_start, date(2026, 4, 1))
        self.assertEqual(combined.period_end, date(2026, 5, 31))
        self.assertEqual(row.name, "Товар новое имя")
        self.assertEqual(row.category, "Новая категория")
        self.assertEqual(row.units, 5)
        self.assertEqual(row.material_sold, 360)
        self.assertEqual(row.labor_sold, 140)
        self.assertEqual(row.material_cost, 72)
        self.assertEqual(row.labor_cost, 28)
        self.assertEqual(row.cost_sold, 500)
        self.assertEqual(row.tax(combined.tax_rate), 74)
        # Ставка для сценария — ставка самого позднего отчета, а не средневзвешенная.
        self.assertEqual(combined.tax_rate, 0.06)
        self.assertEqual(row.net_profit(combined.tax_rate), 326)
        self.assertAlmostEqual(row.profitability(combined.tax_rate), 326 / 500)
        self.assertEqual(combined.unallocated_total, -25)
        self.assertEqual(combined.unallocated["Подписка"], (2, -25))
        self.assertEqual(combined.accrual_stats["Подписка"], (0, 2))
        self.assertEqual(combined.realization_revenue, 150)
        self.assertEqual(combined.realization_units, 3)
        self.assertEqual(combined.totals()["net_profit"], 326)
        self.assertEqual(
            combined.revenue_shares(),
            {
                "commission_share": 0.0,
                "logistics_share": 0.0,
                "points_share": 0.0,
                "net_margin": 301 / 1400,
            },
        )

    def test_rejects_empty_selection(self) -> None:
        with self.assertRaisesRegex(ValueError, "не выбраны"):
            aggregate_calculations([])

    def test_combined_overview_preserves_each_period_unallocated_income_tax(self) -> None:
        april = RunCalculation(
            run_id=1,
            period_start=date(2026, 4, 1),
            period_end=date(2026, 4, 30),
            tax_rate=0.04,
            products=[ProductResult("A", "Товар", 0, 0, revenue_no_points=100)],
            unallocated_total=100,
            unallocated={"Брак по вине Ozon на складе": (1, 100)},
            accrual_stats={},
        )
        may = RunCalculation(
            run_id=2,
            period_start=date(2026, 5, 1),
            period_end=date(2026, 5, 31),
            tax_rate=0.06,
            products=[ProductResult("A", "Товар", 0, 0, revenue_no_points=100)],
            unallocated_total=100,
            unallocated={"Потеря по вине Ozon в логистике": (1, 100)},
            accrual_stats={},
        )

        combined = aggregate_calculations([april, may])

        self.assertEqual(combined.taxable_unallocated_income, 200)
        self.assertEqual(combined.unallocated_income_tax, 10)
        self.assertEqual(combined.totals()["tax"], 20)


if __name__ == "__main__":
    unittest.main()
