"""Налоговые ставки УСН «Доходы» по периодам действия."""

from __future__ import annotations

import math
from bisect import bisect_right
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Iterable


# Ставки считаются равными, если различаются меньше чем на 1e-12 (0,0000000001 %).
RATE_TOLERANCE = 1e-12


class TaxRateError(ValueError):
    """Справочник налоговых ставок заполнен некорректно."""


@dataclass(frozen=True, slots=True)
class TaxRatePeriod:
    """Строка справочника: ставка действует с даты ``valid_from`` (None — «с начала»)."""

    valid_from: date | None
    rate: float


@dataclass(frozen=True, slots=True)
class AppliedTaxRate:
    """Ставка, применённая к отчёту в интервале дат ``start``–``end`` включительно."""

    start: date | None
    end: date | None
    rate: float


class TaxRateSchedule:
    """Упорядоченный справочник ставок; ставка на дату — строка с наибольшей датой ≤ этой даты."""

    __slots__ = ("_periods", "_dates")

    def __init__(self, periods: Iterable[TaxRatePeriod]):
        items = list(periods)
        validate_tax_rate_periods(items)
        self._periods: tuple[TaxRatePeriod, ...] = tuple(
            sorted(items, key=lambda item: (item.valid_from is not None, item.valid_from or date.min))
        )
        self._dates: list[date] = [item.valid_from for item in self._periods[1:]]  # type: ignore[misc]

    @classmethod
    def single(cls, rate: float) -> TaxRateSchedule:
        return cls([TaxRatePeriod(None, float(rate))])

    @classmethod
    def coerce(cls, value: TaxRateSchedule | float | int) -> TaxRateSchedule:
        """Принять справочник или одну ставку (для совместимости со старыми вызовами)."""
        if isinstance(value, TaxRateSchedule):
            return value
        return cls.single(float(value))

    @property
    def periods(self) -> tuple[TaxRatePeriod, ...]:
        return self._periods

    def rate_on(self, day: date) -> float:
        return self._periods[bisect_right(self._dates, day)].rate

    def rate_changes(self, start: date, end: date) -> list[tuple[date, float, float]]:
        """Даты смены ставки в интервале ``start < дата ≤ end``: (дата, прежняя, новая)."""
        changes: list[tuple[date, float, float]] = []
        for index, valid_from in enumerate(self._dates, start=1):
            if not start < valid_from <= end:
                continue
            previous = self._periods[index - 1].rate
            current = self._periods[index].rate
            if not same_rate(previous, current):
                changes.append((valid_from, previous, current))
        return changes

    def applied_rates(self, start: date | None, end: date | None) -> list[AppliedTaxRate]:
        """Ставки справочника в интервале дат; соседние строки с одинаковой ставкой объединяются."""
        if start is None or end is None:
            segments = [
                AppliedTaxRate(
                    item.valid_from,
                    self._dates[index] - timedelta(days=1) if index < len(self._dates) else None,
                    item.rate,
                )
                for index, item in enumerate(self._periods)
            ]
            return _merge_touching(segments)
        if end < start:
            start, end = end, start
        segments: list[AppliedTaxRate] = []
        cursor = start
        while cursor <= end:
            index = bisect_right(self._dates, cursor)
            next_change = self._dates[index] if index < len(self._dates) else None
            segment_end = end if next_change is None else min(end, next_change - timedelta(days=1))
            segments.append(AppliedTaxRate(cursor, segment_end, self._periods[index].rate))
            if segment_end >= end:
                break
            cursor = segment_end + timedelta(days=1)
        return _merge_touching(segments)

    def same_rates_between(
        self,
        other: TaxRateSchedule,
        start: date | None,
        end: date | None,
    ) -> bool:
        """True, если оба справочника дают одинаковую ставку на каждую дату интервала.

        Без интервала сравниваются все даты.
        """
        if start is not None and end is not None and end < start:
            start, end = end, start
        points = {start} if start is not None else {date.min}
        for valid_from in (*self._dates, *other._dates):
            if (start is None or valid_from > start) and (end is None or valid_from <= end):
                points.add(valid_from)
        return all(same_rate(self.rate_on(point), other.rate_on(point)) for point in points)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, TaxRateSchedule):
            return NotImplemented
        return len(self._periods) == len(other._periods) and all(
            left.valid_from == right.valid_from and same_rate(left.rate, right.rate)
            for left, right in zip(self._periods, other._periods)
        )

    def __hash__(self) -> int:
        return hash(tuple((item.valid_from, round(item.rate, 12)) for item in self._periods))

    def __repr__(self) -> str:
        return f"TaxRateSchedule({list(self._periods)!r})"


def validate_tax_rate_periods(periods: list[TaxRatePeriod]) -> None:
    if not periods:
        raise TaxRateError("В справочнике должна быть хотя бы одна налоговая ставка")
    openings = [item for item in periods if item.valid_from is None]
    if len(openings) != 1:
        raise TaxRateError("Первая строка справочника должна действовать «с начала» и быть единственной")
    seen: set[date] = set()
    for item in periods:
        if item.valid_from is not None:
            if item.valid_from in seen:
                raise TaxRateError(
                    f"Дата «{item.valid_from:%d.%m.%Y}» повторяется в справочнике ставок"
                )
            seen.add(item.valid_from)
        if (
            isinstance(item.rate, bool)
            or not isinstance(item.rate, (int, float))
            or not math.isfinite(float(item.rate))
            or item.rate < 0
            or item.rate > 1
        ):
            raise TaxRateError("Налоговая ставка должна быть от 0 до 100%")


def same_rate(left: float, right: float) -> bool:
    return abs(left - right) <= RATE_TOLERANCE


def merge_applied_rates(items: Iterable[AppliedTaxRate]) -> list[AppliedTaxRate]:
    """Объединить ставки нескольких отчётов: соседние и пересекающиеся периоды одной ставки склеиваются."""
    ordered = sorted(
        items,
        key=lambda item: (item.start or date.min, item.end or date.max, item.rate),
    )
    return _merge_touching(ordered)


def format_rate_percent(rate: float) -> str:
    """0.04 → «4 %», 0.065 → «6,5 %»."""
    text = f"{rate * 100:.4f}".rstrip("0").rstrip(".")
    return f"{text.replace('.', ',')} %"


def format_applied_rates(items: Iterable[AppliedTaxRate]) -> str:
    parts: list[str] = []
    for item in items:
        if item.start is None and item.end is None:
            period = "весь период"
        elif item.start is None:
            period = f"по {item.end:%d.%m.%Y}"
        elif item.end is None:
            period = f"с {item.start:%d.%m.%Y}"
        elif item.start == item.end:
            period = f"{item.start:%d.%m.%Y}"
        else:
            period = f"{item.start:%d.%m.%Y}–{item.end:%d.%m.%Y}"
        parts.append(f"{format_rate_percent(item.rate)} — {period}")
    return "; ".join(parts)


def _merge_touching(items: list[AppliedTaxRate]) -> list[AppliedTaxRate]:
    merged: list[AppliedTaxRate] = []
    for item in items:
        if merged:
            last = merged[-1]
            touches = (
                last.end is None
                or item.start is None
                or item.start <= last.end + timedelta(days=1)
            )
            if touches and same_rate(last.rate, item.rate):
                end = None if last.end is None or item.end is None else max(last.end, item.end)
                merged[-1] = AppliedTaxRate(last.start, end, last.rate)
                continue
        merged.append(item)
    return merged
