from __future__ import annotations

from collections import defaultdict
from datetime import date

from .models import ProductResult, RunCalculation
from .tax_rates import merge_applied_rates


SUM_FIELDS = (
    "units",
    "revenue_no_points",
    "partner_programs",
    "points",
    "commission",
    "processing",
    "delivery",
    "logistics",
    "reverse_logistics",
    "returns_cancels",
    "acquiring",
    "stars",
    "packaging",
    "compensation",
    "other",
    "financial_result",
)


def aggregate_calculations(calculations: list[RunCalculation]) -> RunCalculation:
    """Combine saved runs while preserving each run's exact costs and stored taxes."""
    if not calculations:
        raise ValueError("Для обзора не выбраны отчеты")
    ordered = sorted(
        calculations,
        key=lambda item: (
            item.period_start or date.max,
            item.period_end or date.max,
            item.run_id or 0,
        ),
    )
    rows: dict[str, ProductResult] = {}
    material_sold: dict[str, float] = defaultdict(float)
    labor_sold: dict[str, float] = defaultdict(float)
    taxes: dict[str, float] = defaultdict(float)

    for calculation in ordered:
        for source in calculation.products:
            target = rows.get(source.article)
            if target is None:
                target = ProductResult(
                    article=source.article,
                    name=source.name,
                    category=source.category,
                    material_cost=source.material_cost,
                    labor_cost=source.labor_cost,
                )
                rows[source.article] = target
            else:
                # The newest saved report supplies descriptive and fallback unit data.
                target.name = source.name or target.name
                target.category = source.category
                target.material_cost = source.material_cost
                target.labor_cost = source.labor_cost
            for field_name in SUM_FIELDS:
                setattr(
                    target,
                    field_name,
                    float(getattr(target, field_name)) + float(getattr(source, field_name)),
                )
            material_sold[source.article] += source.material_sold
            labor_sold[source.article] += source.labor_sold
            taxes[source.article] += source.tax(calculation.tax_rate)

    for article, target in rows.items():
        target.material_sold_override = material_sold[article]
        target.labor_sold_override = labor_sold[article]
        target.tax_override = taxes[article]
        if target.units:
            target.material_cost = material_sold[article] / target.units
            target.labor_cost = labor_sold[article] / target.units

    unallocated: dict[str, list[float]] = {}
    accrual_stats: dict[str, list[int]] = {}
    for calculation in ordered:
        for accrual_type, (row_count, amount) in calculation.unallocated.items():
            values = unallocated.setdefault(accrual_type, [0.0, 0.0])
            values[0] += row_count
            values[1] += amount
        for accrual_type, (with_article, without_article) in calculation.accrual_stats.items():
            values = accrual_stats.setdefault(accrual_type, [0, 0])
            values[0] += with_article
            values[1] += without_article

    starts = [item.period_start for item in ordered if item.period_start is not None]
    ends = [item.period_end for item in ordered if item.period_end is not None]
    tax_starts = [item.tax_period_start for item in ordered if item.tax_period_start is not None]
    tax_ends = [item.tax_period_end for item in ordered if item.tax_period_end is not None]
    # Сценарий цены использует ставку самого позднего отчета на дату окончания его периода.
    latest = max(
        ordered,
        key=lambda item: (
            item.period_end or date.min,
            item.period_start or date.min,
            item.run_id or 0,
        ),
    )
    return RunCalculation(
        run_id=None,
        period_start=min(starts) if starts else None,
        period_end=max(ends) if ends else None,
        tax_rate=latest.tax_rate,
        products=sorted(rows.values(), key=lambda item: item.article.casefold()),
        unallocated_total=sum(item.unallocated_total for item in ordered),
        unallocated={
            name: (int(values[0]), float(values[1]))
            for name, values in unallocated.items()
        },
        accrual_stats={
            name: (int(values[0]), int(values[1]))
            for name, values in accrual_stats.items()
        },
        skipped_articles={
            article: message
            for calculation in ordered
            for article, message in calculation.skipped_articles.items()
        },
        sku_conflicts=set().union(*(item.sku_conflicts for item in ordered)),
        duplicate_realization_rows=sum(item.duplicate_realization_rows for item in ordered),
        already_accrued_realization_rows=sum(
            item.already_accrued_realization_rows for item in ordered
        ),
        realization_revenue=sum(item.realization_revenue for item in ordered),
        realization_units=sum(item.realization_units for item in ordered),
        source_period_warnings=[
            warning
            for calculation in ordered
            for warning in calculation.source_period_warnings
        ],
        taxable_unallocated_income_override=sum(
            item.taxable_unallocated_income for item in ordered
        ),
        double_count_warnings=[
            warning
            for calculation in ordered
            for warning in calculation.double_count_warnings
        ],
        # Налог каждого отчета уже рассчитан по ставкам на даты строк: только суммируем.
        unallocated_income_tax_override=sum(
            item.unallocated_income_tax for item in ordered
        ),
        tax_period_start=min(tax_starts) if tax_starts else None,
        tax_period_end=max(tax_ends) if tax_ends else None,
        applied_tax_rates=merge_applied_rates(
            rate for calculation in ordered for rate in calculation.applied_tax_rates
        ),
        tax_rate_warnings=[
            warning
            for calculation in ordered
            for warning in calculation.tax_rate_warnings
        ],
    )
