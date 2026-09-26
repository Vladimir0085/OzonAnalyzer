from __future__ import annotations

from collections import defaultdict
from dataclasses import replace
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from .excel_reader import all_additional_income_rows, all_rows, normalize_text
from .models import AccrualRow, ParsedSource, Product, ProductResult, RunCalculation, ScenarioRow, UnknownProduct
from .tax_rates import TaxRateSchedule, format_rate_percent


class CalculationError(ValueError):
    pass


def _type_key(value: str) -> str:
    return normalize_text(value).replace("ё", "е")


_SALE_COMPONENT_TYPES = {
    _type_key("Выручка"),
    _type_key("Программы партнёров"),
    _type_key("Баллы за скидки"),
}
_SALES_SERVICE_GROUP = _type_key("Продажи")
# Строки товара, входящие в налоговую базу УСН «Доходы»; баллы в базу не входят.
_TAXABLE_PRODUCT_FIELDS = {"revenue_no_points", "partner_programs"}


def _is_sale_component(row: AccrualRow, type_key: str) -> bool:
    """Return True only for components that belong to an actual sale.

    Ozon repeats partner-program and points rows in the ``Возвраты``
    service group with negative amounts.  Treating those rows as sales and then
    subtracting ``Возврат выручки`` quantity counts the same return twice.
    Older saved reports may not expose ``Группа услуг``; for them a positive
    amount is the safest compatible indication of a sale component.
    """
    if type_key not in _SALE_COMPONENT_TYPES:
        return False
    group_key = _type_key(row.service_group)
    if group_key:
        return group_key == _SALES_SERVICE_GROUP
    return row.amount > 0


def distribution_status(with_article: int, without_article: int, empty: str = "НЕТ ДАННЫХ") -> str:
    if with_article > 0 and without_article > 0:
        return "СМЕШАННОЕ РАСПРЕДЕЛЕНИЕ"
    if with_article > 0:
        return "ПО АРТИКУЛУ"
    if without_article > 0:
        return "НЕРАСПРЕДЕЛЕННЫЕ ДОХОДЫ/РАСХОДЫ"
    return empty


def accrual_category(accrual_type: str) -> tuple[str, str]:
    key = _type_key(accrual_type)
    if key in {_type_key("Выручка"), _type_key("Возврат выручки")}:
        return "revenue_no_points", "Выручка без баллов"
    if key == _type_key("Программы партнёров"):
        return "partner_programs", "Программы партнёров"
    if key == _type_key("Баллы за скидки"):
        return "points", "Баллы за скидки"
    if key in {_type_key("Вознаграждение за продажу"), _type_key("Возврат вознаграждения")}:
        return "commission", "Комиссия Ozon"
    if key.startswith(_type_key("Обработка отправления")):
        return "processing", "Обработка отправления Drop-off"
    if key.startswith(_type_key("Доставка до места выдачи")):
        return "delivery", "Доставка до места выдачи"
    if key == _type_key("Логистика"):
        return "logistics", "Логистика"
    if key == _type_key("Обратная логистика"):
        return "reverse_logistics", "Обратная логистика"
    if (
        key.startswith(_type_key("Обработка возвратов"))
        or key.startswith(_type_key("Обработка возврата"))
        or key in {
            _type_key("Обработка отменённых и невостребованных товаров"),
            _type_key("Обработка частичного невыкупа"),
        }
    ):
        return "returns_cancels", "Возвраты, отмены и невыкупы"
    if key == _type_key("Эквайринг"):
        return "acquiring", "Эквайринг"
    if key == _type_key("Звёздные товары"):
        return "stars", "Продвижение «Звёздные товары»"
    if key in {
        _type_key("Упаковка товара партнёрами"),
        _type_key("Обеспечение материалами для упаковки товара"),
    }:
        return "packaging", "Упаковка и материалы"
    if key in {
        _type_key("Потеря по вине Ozon в логистике"),
        _type_key("Брак по вине Ozon на складе"),
    }:
        return "compensation", "Компенсации Ozon"
    return "other", "Прочие начисления по товару"


def guide_target(accrual_type: str) -> str:
    return f"С артикулом → {accrual_category(accrual_type)[1]}; без артикула → нераспределенные"


def build_sku_map(accrual_rows, allowed_articles: set[str] | None = None) -> tuple[dict[str, str], set[str]]:
    mapping: dict[str, str] = {}
    conflicts: set[str] = set()
    for row in accrual_rows:
        if not row.sku or not row.article:
            continue
        if allowed_articles is not None and row.article not in allowed_articles:
            continue
        if row.sku in conflicts:
            continue
        existing = mapping.get(row.sku)
        if existing and existing != row.article:
            mapping.pop(row.sku, None)
            conflicts.add(row.sku)
        else:
            mapping[row.sku] = row.article
    return mapping, conflicts


def discover_unknown_products(sources: list[ParsedSource], products: dict[str, Product]) -> list[UnknownProduct]:
    accrual_rows, realization_rows = all_rows(sources)
    provisional_map, conflicts = build_sku_map(accrual_rows)
    unknown: dict[str, UnknownProduct] = {}

    def register(article: str, name: str, sku: str, source_name: str) -> None:
        if not article or article in products:
            return
        item = unknown.setdefault(article, UnknownProduct(article=article, name=name or article, sku=sku))
        if item.name == item.article and name:
            item.name = name
        if not item.sku and sku:
            item.sku = sku
        item.source_names.add(source_name)

    for row in accrual_rows:
        register(row.article, row.product_name, row.sku, row.source_name)

    known_or_pending = set(products) | set(unknown)
    for row in realization_rows:
        mapped = provisional_map.get(row.sku, "") if row.sku not in conflicts else ""
        if mapped and mapped in known_or_pending:
            continue
        if row.raw_article in products:
            continue
        register(row.raw_article, row.product_name, row.sku, row.source_name)
    return sorted(unknown.values(), key=lambda item: item.article.casefold())


def calculate_run(
    sources: list[ParsedSource],
    products: dict[str, Product],
    tax_rates: TaxRateSchedule | float,
    skipped_articles: set[str] | None = None,
    *,
    today: date | None = None,
) -> RunCalculation:
    """Рассчитать отчёт; налог каждой строки — по ставке справочника на дату этой строки."""
    schedule = TaxRateSchedule.coerce(tax_rates)
    skipped_articles = skipped_articles or set()
    accrual_rows, realization_rows = all_rows(sources)
    additional_income_rows = all_additional_income_rows(sources)
    referenced_articles = {
        row.article for row in accrual_rows if row.article
    } | {
        row.raw_article for row in realization_rows if row.raw_article
    }
    # An archived catalog item can still have a legitimate charge (for example,
    # acquiring, logistics or return processing).  Its rows must not disappear
    # merely because there were no sales or the item was archived after the sale.
    calculation_products = {
        key: value
        for key, value in products.items()
        if (value.active or key in referenced_articles) and key not in skipped_articles
    }
    results = {
        article: ProductResult(
            article=product.article,
            name=product.name,
            material_cost=product.material_cost,
            labor_cost=product.labor_cost,
            category=product.category,
        )
        for article, product in calculation_products.items()
    }
    sku_map, sku_conflicts = build_sku_map(accrual_rows, set(calculation_products))
    current_stats: dict[str, list[object]] = {}
    breakdown: dict[str, list[object]] = {}
    skipped_detail: dict[str, str] = {}
    skipped_accrual_amounts: dict[str, float] = defaultdict(float)
    skipped_realization_amounts: dict[str, float] = defaultdict(float)
    sales_orders: dict[tuple[str, str, float], list[float]] = {}
    returns_by_article: dict[str, float] = defaultdict(float)
    accrual_revenue_keys: set[tuple[str, str]] = set()
    unallocated_total = 0.0
    taxable_unallocated_income = 0.0
    allocated_accrual_total = 0.0
    skipped_accrual_total = 0.0

    dates = [row.accrual_date for row in accrual_rows if row.accrual_date]
    period_start: date | None = min(dates) if dates else None
    period_end: date | None = max(dates) if dates else None

    # Строка без собственной даты облагается по ставке на дату окончания периода отчёта.
    # Если в отчёте вообще нет дат, остаётся только ставка на сегодня.
    fallback_tax_date = period_end or today or date.today()
    tax_rate_warnings: list[str] = []
    tax_dates: list[date] = []
    product_tax_bases: dict[str, dict[float, float]] = defaultdict(lambda: defaultdict(float))
    unallocated_tax_bases: dict[float, float] = defaultdict(float)
    fallback_used = False

    def tax_base_rate(row_date: date | None) -> float:
        nonlocal fallback_used
        if row_date is None:
            fallback_used = True
        tax_date = row_date or fallback_tax_date
        tax_dates.append(tax_date)
        return schedule.rate_on(tax_date)

    for row in accrual_rows:
        type_key = _type_key(row.accrual_type)
        stat = current_stats.setdefault(type_key, [row.accrual_type, 0, 0])
        stat[1 if row.article else 2] = int(stat[1 if row.article else 2]) + 1

        if row.article:
            if row.article not in results or row.article in skipped_articles:
                skipped_detail.setdefault(
                    row.article,
                    f"{row.article} — {row.product_name or row.article} ({row.source_name}, строка {row.row_number})",
                )
                skipped_accrual_amounts[row.article] += row.amount
                skipped_accrual_total += row.amount
                continue
            result = results[row.article]
            result.financial_result += row.amount
            allocated_accrual_total += row.amount
            field_name, _ = accrual_category(row.accrual_type)
            setattr(result, field_name, getattr(result, field_name) + row.amount)
            if field_name in _TAXABLE_PRODUCT_FIELDS:
                product_tax_bases[row.article][tax_base_rate(row.accrual_date)] += row.amount

            if row.sku and row.sku not in sku_conflicts:
                existing = sku_map.get(row.sku)
                if existing and existing != row.article:
                    sku_map.pop(row.sku, None)
                    sku_conflicts.add(row.sku)
                elif not existing:
                    sku_map[row.sku] = row.article

            key = _type_key(row.accrual_type)
            if _is_sale_component(row, key):
                if row.accrual_id and row.sku:
                    accrual_revenue_keys.add((row.sku, normalize_text(row.accrual_id)))
                order_id = row.accrual_id or f"{row.source_name}#{row.row_number}"
                seller_price = round(abs(row.seller_price), 6)
                order_key = (order_id, row.article, seller_price)
                order = sales_orders.setdefault(order_key, [0.0, 0.0])
                order[0] += row.amount
                if key == _type_key("Выручка"):
                    order[1] += abs(row.quantity)
            elif key == _type_key("Возврат выручки"):
                returns_by_article[row.article] += abs(row.quantity)
        else:
            unallocated_total += row.amount
            if row.amount > 0:
                taxable_unallocated_income += row.amount
                unallocated_tax_bases[tax_base_rate(row.accrual_date)] += row.amount
            item = breakdown.setdefault(type_key or "без типа начисления", [row.accrual_type or "Без типа начисления", 0, 0.0])
            item[1] = int(item[1]) + 1
            item[2] = float(item[2]) + row.amount

    for row in additional_income_rows:
        type_key = _type_key(row.income_type)
        stat = current_stats.setdefault(type_key, [row.income_type, 0, 0])
        stat[2] = int(stat[2]) + 1
        unallocated_total += row.amount
        if row.amount > 0:
            taxable_unallocated_income += row.amount
            unallocated_tax_bases[tax_base_rate(row.income_date)] += row.amount
        item = breakdown.setdefault(type_key, [row.income_type, 0, 0.0])
        item[1] = int(item[1]) + 1
        item[2] = float(item[2]) + row.amount

    for (_, article, seller_price), (amount, source_quantity) in sales_orders.items():
        if article not in results:
            continue
        if seller_price > 0:
            units = _round_half_up(abs(amount) / seller_price)
        else:
            units = source_quantity
        results[article].units += units

    for article, quantity in returns_by_article.items():
        if article in results:
            results[article].units -= quantity

    realization_source_by_row = {
        id(row): source for source in sources for row in source.realization_rows
    }
    undated_realization_sources: dict[int, ParsedSource] = {}
    realization_keys: dict[tuple[str, str], tuple[float, float, str]] = {}
    duplicate_rows = 0
    already_accrued_rows = 0
    realization_revenue = 0.0
    realization_units = 0.0
    for row in realization_rows:
        shipment = normalize_text(row.shipment) or f"{row.source_name}#{row.row_number}"
        row_key = (row.sku, shipment)
        if row_key in accrual_revenue_keys:
            already_accrued_rows += 1
            continue
        if row_key in realization_keys:
            old_amount, old_quantity, _ = realization_keys[row_key]
            if abs(old_amount - row.amount) <= 0.05 and abs(old_quantity - row.quantity) <= 0.0001:
                duplicate_rows += 1
                continue
            raise CalculationError(
                f"Противоречивый дубль выкупленного товара для отправления {row.shipment}."
            )
        realization_keys[row_key] = (row.amount, row.quantity, row.source_name)

        article = ""
        if row.sku and row.sku not in sku_conflicts:
            mapped = sku_map.get(row.sku, "")
            if mapped in results:
                article = mapped
        if not article and row.raw_article in results:
            article = row.raw_article
        if not article or article in skipped_articles:
            candidate = (
                article
                or row.raw_article
                or (f"SKU {row.sku}" if row.sku else "Неопознанный товар")
            )
            skipped_detail.setdefault(
                candidate,
                f"{candidate} — {row.product_name or candidate} ({row.source_name}, строка {row.row_number})",
            )
            skipped_realization_amounts[candidate] += row.amount
            continue
        result = results[article]
        result.units += row.quantity
        result.revenue_no_points += row.amount
        result.financial_result += row.amount
        realization_revenue += row.amount
        realization_units += row.quantity
        realization_source = realization_source_by_row.get(id(row))
        tax_date = row.sale_date
        if tax_date is None:
            if realization_source is not None:
                undated_realization_sources[id(realization_source)] = realization_source
                tax_date = realization_source.period_end
        product_tax_bases[article][tax_base_rate(tax_date)] += row.amount

    source_accrual_total = sum(row.amount for row in accrual_rows) + sum(
        row.amount for row in additional_income_rows
    )
    allocation_difference = (
        source_accrual_total
        - allocated_accrual_total
        - unallocated_total
        - skipped_accrual_total
    )
    if abs(allocation_difference) > 0.01:
        raise CalculationError(
            "Не сошелся контроль распределения начислений: "
            f"расхождение {allocation_difference:.2f} руб."
        )

    for article, message in list(skipped_detail.items()):
        details: list[str] = []
        accrual_amount = skipped_accrual_amounts.get(article, 0.0)
        realization_amount = skipped_realization_amounts.get(article, 0.0)
        if abs(accrual_amount) > 0.000001:
            details.append(f"начисления: {_money_ru(accrual_amount)}")
        if abs(realization_amount) > 0.000001:
            details.append(f"выкупленная выручка: {_money_ru(realization_amount)}")
        if details:
            skipped_detail[article] = f"{message}; пропущенная сумма — {', '.join(details)}"

    for source in undated_realization_sources.values():
        start = source.period_start or period_start
        end = source.period_end or period_end
        if start is None or end is None:
            continue
        changes = schedule.rate_changes(start, end)
        if not changes:
            continue
        tax_date = source.period_end or fallback_tax_date
        change_text = ", ".join(
            f"с {changed:%d.%m.%Y}: {format_rate_percent(old)} → {format_rate_percent(new)}"
            for changed, old, new in changes
        )
        tax_rate_warnings.append(
            f"«{source.path.name}»: период выкупов {start:%d.%m.%Y}–{end:%d.%m.%Y} "
            f"пересекает смену налоговой ставки ({change_text}). В файле нет дат строк, "
            f"поэтому налог с выкупов рассчитан по ставке на {tax_date:%d.%m.%Y} — "
            f"{format_rate_percent(schedule.rate_on(tax_date))}."
        )
    if period_end is None and fallback_used:
        tax_rate_warnings.append(
            "В отчёте по начислениям нет дат, поэтому налог рассчитан по ставке на "
            f"{fallback_tax_date:%d.%m.%Y} — {format_rate_percent(schedule.rate_on(fallback_tax_date))}."
        )

    for article, bases in product_tax_bases.items():
        if article in results:
            results[article].tax_override = sum(base * rate for rate, base in bases.items())
    for result in results.values():
        if result.tax_override is None:
            result.tax_override = 0.0
    unallocated_income_tax = sum(base * rate for rate, base in unallocated_tax_bases.items())

    source_dates = [
        value
        for source in sources
        for value in (source.period_start, source.period_end)
        if value is not None
    ]
    period_dates = [value for value in (period_start, period_end) if value is not None]
    all_tax_dates = tax_dates + source_dates + period_dates
    tax_period_start = min(all_tax_dates) if all_tax_dates else None
    tax_period_end = max(all_tax_dates) if all_tax_dates else None

    stats = {
        str(values[0]): (int(values[1]), int(values[2]))
        for values in sorted(current_stats.values(), key=lambda value: str(value[0]).casefold())
    }
    unallocated = {
        str(values[0]): (int(values[1]), float(values[2]))
        for values in sorted(breakdown.values(), key=lambda value: str(value[0]).casefold())
    }
    return RunCalculation(
        run_id=None,
        period_start=period_start,
        period_end=period_end,
        tax_rate=schedule.rate_on(fallback_tax_date),
        products=sorted(results.values(), key=lambda item: item.article.casefold()),
        unallocated_total=unallocated_total,
        unallocated=unallocated,
        accrual_stats=stats,
        source_files=sources,
        skipped_articles=skipped_detail,
        sku_conflicts=sku_conflicts,
        duplicate_realization_rows=duplicate_rows,
        already_accrued_realization_rows=already_accrued_rows,
        realization_revenue=realization_revenue,
        realization_units=realization_units,
        taxable_unallocated_income_override=taxable_unallocated_income,
        unallocated_income_tax_override=unallocated_income_tax,
        tax_schedule=schedule,
        tax_period_start=tax_period_start,
        tax_period_end=tax_period_end,
        applied_tax_rates=(
            schedule.applied_rates(tax_period_start, tax_period_end)
            if tax_period_start is not None
            else []
        ),
        tax_rate_warnings=tax_rate_warnings,
    )


def calculate_scenario(result: ProductResult, tax_rate: float, planned_price: float | None = None) -> ScenarioRow:
    current_price = result.average_price()
    if planned_price is None:
        planned_price = current_price
    if not result.units or planned_price is None:
        return ScenarioRow(
            article=result.article,
            name=result.name,
            category=result.category,
            unit_cost=result.total_cost,
            units=result.units,
            current_price=current_price,
            planned_price=planned_price,
            price_change=None,
            profitability=None,
            ozon_costs_without_commission=None,
            planned_revenue=None,
            commission_rate=None,
            planned_commission=None,
            planned_points=None,
            taxable_base=None,
            tax=None,
            profit=None,
            profit_per_unit_before_cost=None,
            net_profit_per_unit=None,
        )
    price_change = planned_price / current_price - 1 if current_price else None
    ozon_costs_without_commission = (
        result.revenue_including_points - result.financial_result + result.commission
    )
    planned_revenue = planned_price * result.units
    commission_rate = -result.commission / result.revenue_including_points if result.revenue_including_points else 0.0
    planned_commission = -planned_revenue * commission_rate
    planned_points = (
        result.points / result.revenue_including_points * planned_revenue
        if result.revenue_including_points
        else 0.0
    )
    taxable_base = planned_revenue - planned_points
    tax = taxable_base * tax_rate
    profit = planned_revenue - ozon_costs_without_commission + planned_commission - tax
    profit_per_unit_before_cost = profit / result.units
    net_profit_per_unit = profit_per_unit_before_cost - result.total_cost
    profitability = net_profit_per_unit / result.total_cost if result.total_cost else 0.0
    return ScenarioRow(
        article=result.article,
        name=result.name,
        category=result.category,
        unit_cost=result.total_cost,
        units=result.units,
        current_price=current_price,
        planned_price=planned_price,
        price_change=price_change,
        profitability=profitability,
        ozon_costs_without_commission=ozon_costs_without_commission,
        planned_revenue=planned_revenue,
        commission_rate=commission_rate,
        planned_commission=planned_commission,
        planned_points=planned_points,
        taxable_base=taxable_base,
        tax=tax,
        profit=profit,
        profit_per_unit_before_cost=profit_per_unit_before_cost,
        net_profit_per_unit=net_profit_per_unit,
    )


def _round_half_up(value: float) -> float:
    return float(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _money_ru(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ").replace(".", ",") + " руб."


def clone_result(result: ProductResult) -> ProductResult:
    return replace(result)
