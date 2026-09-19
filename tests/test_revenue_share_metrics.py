from __future__ import annotations

import unittest

from ozon_app.models import ProductResult, RunCalculation
from ozon_app.report_totals import overview_revenue_kpi_values
from ozon_app.ui import _result_values


class RevenueShareMetricTests(unittest.TestCase):
    def test_august_logistics_amount_includes_all_four_report_columns(self) -> None:
        product = ProductResult(
            article="AUGUST",
            name="Август 2026",
            material_cost=0,
            labor_cost=0,
            delivery=-1_378.85,
            logistics=-5_526.49,
            reverse_logistics=-309.08,
            returns_cancels=0,
        )
        calculation = RunCalculation(
            run_id=1,
            period_start=None,
            period_end=None,
            tax_rate=0.06,
            products=[product],
            unallocated_total=0,
            unallocated={},
            accrual_stats={},
        )

        self.assertAlmostEqual(calculation.revenue_amounts()["logistics"], 7_214.42)

    def test_product_and_report_shares_use_revenue_including_points(self) -> None:
        product = ProductResult(
            article="A",
            name="Товар",
            material_cost=50,
            labor_cost=50,
            units=2,
            revenue_no_points=800,
            partner_programs=100,
            points=100,
            commission=-200,
            delivery=-30,
            logistics=-80,
            reverse_logistics=-20,
            returns_cancels=-10,
            financial_result=600,
        )
        calculation = RunCalculation(
            run_id=1,
            period_start=None,
            period_end=None,
            tax_rate=0.04,
            products=[product],
            unallocated_total=-50,
            unallocated={"Общие расходы": (1, -50)},
            accrual_stats={},
        )

        self.assertAlmostEqual(product.commission_share(), 0.2)
        self.assertEqual(product.logistics_total, -140)
        self.assertAlmostEqual(product.logistics_share(), 0.14)
        self.assertAlmostEqual(product.points_share(), 0.1)
        self.assertAlmostEqual(product.net_margin(0.04), 0.364)
        self.assertEqual(
            calculation.revenue_shares(),
            {
                "commission_share": 0.2,
                "logistics_share": 0.14,
                "points_share": 0.1,
                "net_margin": 0.314,
            },
        )
        self.assertEqual(
            calculation.revenue_amounts(),
            {"commission": 200, "logistics": 140, "points": 100},
        )
        self.assertEqual(
            overview_revenue_kpi_values(calculation),
            {
                "commission": "200.00 ₽ · 20.00%",
                "logistics": "140.00 ₽ · 14.00%",
                "points": "100.00 ₽ · 10.00%",
                "net_margin": "31.40%",
            },
        )
        self.assertEqual(
            _result_values(product, 0.04)[-4:],
            ("20.00%", "14.00%", "10.00%", "36.40%"),
        )

    def test_zero_revenue_returns_zero_shares(self) -> None:
        product = ProductResult("A", "Товар", 0, 0, commission=-10, logistics=-5)

        self.assertEqual(product.commission_share(), 0.0)
        self.assertEqual(product.logistics_share(), 0.0)
        self.assertEqual(product.points_share(), 0.0)
        self.assertEqual(product.net_margin(0.04), 0.0)

    def test_report_net_margin_includes_unallocated_income_and_expenses(self) -> None:
        product = ProductResult(
            "A",
            "Товар",
            100,
            0,
            units=1,
            revenue_no_points=500,
            financial_result=300,
        )
        calculation = RunCalculation(
            run_id=1,
            period_start=None,
            period_end=None,
            tax_rate=0.04,
            products=[product],
            unallocated_total=-80,
            unallocated={"Общие расходы": (1, -80)},
            accrual_stats={},
        )

        self.assertAlmostEqual(product.net_margin(0.04), 0.36)
        self.assertAlmostEqual(calculation.revenue_shares()["net_margin"], 0.20)


if __name__ == "__main__":
    unittest.main()
