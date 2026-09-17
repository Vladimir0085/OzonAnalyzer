from __future__ import annotations

import unittest

from ozon_app.models import ProductResult, RunCalculation
from ozon_app.report_totals import report_total_profitability, report_total_value


class UnallocatedIncomeTaxTests(unittest.TestCase):
    def test_positive_unallocated_income_is_taxable_but_not_product_revenue_or_profit(self) -> None:
        product = ProductResult(
            article="A",
            name="Товар",
            material_cost=100,
            labor_cost=0,
            units=1,
            revenue_no_points=1_000,
            financial_result=600,
        )
        calculation = RunCalculation(
            run_id=1,
            period_start=None,
            period_end=None,
            tax_rate=0.04,
            products=[product],
            unallocated_total=140.78,
            unallocated={
                "Брак по вине Ozon на складе": (1, 120.78),
                "Премия Ozon": (1, 40.0),
                "Подписка": (1, -20.0),
            },
            accrual_stats={},
        )

        self.assertEqual(product.revenue_including_points, 1_000)
        self.assertEqual(product.taxable_income, 1_000)
        self.assertEqual(product.net_profit(0.04), 460)
        self.assertAlmostEqual(calculation.taxable_unallocated_income, 160.78)
        self.assertAlmostEqual(calculation.unallocated_income_tax, 6.4312)
        self.assertAlmostEqual(calculation.totals()["taxable_income"], 1_160.78)
        self.assertAlmostEqual(calculation.totals()["tax"], 46.4312)
        self.assertAlmostEqual(calculation.report_net_profit, 594.3488)
        self.assertAlmostEqual(report_total_value(calculation), 594.3488)
        self.assertAlmostEqual(report_total_profitability(calculation), 5.943488)
        self.assertAlmostEqual(calculation.revenue_shares()["net_margin"], 0.5943488)

    def test_report_total_profitability_is_zero_without_cost_of_goods_sold(self) -> None:
        calculation = RunCalculation(
            run_id=1,
            period_start=None,
            period_end=None,
            tax_rate=0.04,
            products=[],
            unallocated_total=100,
            unallocated={"Премия Ozon": (1, 100)},
            accrual_stats={},
        )

        self.assertEqual(report_total_profitability(calculation), 0.0)


if __name__ == "__main__":
    unittest.main()
