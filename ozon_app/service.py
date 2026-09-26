from __future__ import annotations

import re
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .backup import create_backup
from .calculator import calculate_run, discover_unknown_products
from .config import ensure_app_dirs
from .database import Database, RunReplacement
from .double_count import DoubleCountMatch, find_double_count_matches
from .excel_reader import (
    REPORT_ACCRUAL,
    REPORT_ADDITIONAL_INCOME,
    REPORT_REALIZATION,
    parse_report,
)
from .models import ParsedSource, Product, RunCalculation, RunSummary, UnknownProduct
from .tax_rates import TaxRateSchedule


# Изменение налога меньше полкопейки считается отсутствием изменения.
TAX_CHANGE_TOLERANCE = 0.005


@dataclass(slots=True)
class ImportSession:
    sources: list[ParsedSource]
    unknown_products: list[UnknownProduct]
    duplicate_sources: list[ParsedSource] = field(default_factory=list)

    @property
    def has_accrual(self) -> bool:
        return any(source.report_type == REPORT_ACCRUAL for source in self.sources)

    @property
    def has_realization(self) -> bool:
        return any(source.report_type == REPORT_REALIZATION for source in self.sources)

    @property
    def has_additional_income(self) -> bool:
        return any(
            source.report_type == REPORT_ADDITIONAL_INCOME for source in self.sources
        )

    @property
    def period_start(self) -> date | None:
        dates = [
            row.accrual_date
            for source in self.sources
            for row in source.accrual_rows
            if row.accrual_date is not None
        ]
        return min(dates) if dates else None

    @property
    def period_end(self) -> date | None:
        dates = [
            row.accrual_date
            for source in self.sources
            for row in source.accrual_rows
            if row.accrual_date is not None
        ]
        return max(dates) if dates else None

    def realization_period_warnings(self) -> list[str]:
        realization_sources = [
            source for source in self.sources if source.report_type == REPORT_REALIZATION
        ]
        if not realization_sources:
            return []

        accrual_months = {
            (row.accrual_date.year, row.accrual_date.month)
            for source in self.sources
            for row in source.accrual_rows
            if row.accrual_date is not None
        }
        if not accrual_months:
            return [
                "В отчете по начислениям не удалось определить месяц, "
                "поэтому совместимость периодов не проверена."
            ]

        expected = _month_list(accrual_months)
        warnings: list[str] = []
        for source in realization_sources:
            if source.period_start is None or source.period_end is None:
                warnings.append(
                    f"«{source.path.name}»: период выкупов не удалось определить; "
                    f"ожидаемый месяц: {expected}."
                )
                continue
            start_month = (source.period_start.year, source.period_start.month)
            end_month = (source.period_end.year, source.period_end.month)
            actual_period = (
                f"{source.period_start:%d.%m.%Y}–{source.period_end:%d.%m.%Y}"
            )
            if start_month != end_month:
                warnings.append(
                    f"«{source.path.name}»: период {actual_period} охватывает "
                    f"несколько месяцев; ожидаемый месяц: {expected}."
                )
            elif start_month not in accrual_months:
                warnings.append(
                    f"«{source.path.name}»: период выкупов {actual_period} не относится "
                    f"к месяцу отчета по начислениям ({expected})."
                )
        return warnings

    def double_count_matches(self) -> list[DoubleCountMatch]:
        return find_double_count_matches(self.sources)

    def exclude_source(self, source: ParsedSource) -> None:
        """Drop a supplementary file that the user decided not to import."""
        if source.report_type == REPORT_ACCRUAL:
            raise ValueError("Отчет по начислениям нельзя исключить из расчета")
        self.sources = [item for item in self.sources if item is not source]
        self.duplicate_sources = [
            item for item in self.duplicate_sources if item is not source
        ]


@dataclass(slots=True)
class ImportBatch:
    sessions: list[ImportSession]
    unknown_products: list[UnknownProduct]

    @property
    def sources(self) -> list[ParsedSource]:
        return [source for session in self.sessions for source in session.sources]

    @property
    def duplicate_sources(self) -> list[ParsedSource]:
        return [
            source
            for session in self.sessions
            for source in session.duplicate_sources
        ]


@dataclass(slots=True)
class HistoryRecalculationResult:
    replaced_runs: int
    recovered_product_rows: int
    financial_result_delta: float
    skipped_articles: int
    units_delta: float = 0.0
    net_profit_delta: float = 0.0
    old_to_new: dict[int, int] = field(default_factory=dict)


@dataclass(slots=True)
class TaxRateChangeItem:
    """Отчет, которого касается изменение справочника ставок."""

    run_id: int
    report_name: str
    period_start: date | None
    period_end: date | None
    old_tax: float
    new_tax: float
    old_product_net_profit: float
    new_product_net_profit: float
    old_activity_net_profit: float
    new_activity_net_profit: float

    @property
    def tax_delta(self) -> float:
        return self.new_tax - self.old_tax

    @property
    def product_net_profit_delta(self) -> float:
        return self.new_product_net_profit - self.old_product_net_profit

    @property
    def activity_net_profit_delta(self) -> float:
        return self.new_activity_net_profit - self.old_activity_net_profit

    @property
    def tax_changed(self) -> bool:
        return abs(self.tax_delta) >= TAX_CHANGE_TOLERANCE


@dataclass(slots=True)
class TaxRateChangePlan:
    """Предпросмотр изменения справочника ставок; применяется только после подтверждения."""

    old_schedule: TaxRateSchedule
    new_schedule: TaxRateSchedule
    items: list[TaxRateChangeItem] = field(default_factory=list)
    replacements: list[RunReplacement] = field(default_factory=list)

    @property
    def changed_items(self) -> list[TaxRateChangeItem]:
        """Отчеты, у которых меняется налог (показываются в предпросмотре)."""
        return [item for item in self.items if item.tax_changed]

    @property
    def unchanged_tax_count(self) -> int:
        """Отчеты, у которых обновится только снимок ставок и ставка для сценария."""
        return sum(1 for item in self.items if not item.tax_changed)


class AppService:
    def __init__(self, base_dir: Path | None = None):
        self.paths = ensure_app_dirs(base_dir)
        self.db = Database(self.paths["database"])

    def prepare_import(self, file_paths: list[str | Path]) -> ImportBatch:
        if not file_paths:
            raise ValueError("Не выбраны исходные файлы")
        sources: list[ParsedSource] = []
        duplicates: list[ParsedSource] = []
        selected_hashes: set[str] = set()
        for path in file_paths:
            source = parse_report(path)
            if source.file_hash in selected_hashes:
                raise ValueError(
                    f"Файл «{source.path.name}» выбран повторно. "
                    "Импорт прерван, чтобы не удваивать суммы."
                )
            source.duplicate_run_ids = self.db.find_runs_by_hash(source.file_hash)
            if source.duplicate_run_ids:
                duplicates.append(source)
            selected_hashes.add(source.file_hash)
            sources.append(source)
        if not any(source.report_type == REPORT_ACCRUAL for source in sources):
            raise ValueError("Для расчета нужен хотя бы один отчет по начислениям")
        sessions = split_import_sources(sources)
        products = self.db.product_map(active_only=False)
        unknown = discover_unknown_products(sources, products)
        duplicate_hashes = {source.file_hash for source in duplicates}
        for session in sessions:
            session.duplicate_sources = [
                source for source in session.sources if source.file_hash in duplicate_hashes
            ]
        return ImportBatch(sessions=sessions, unknown_products=unknown)

    def replacement_run_ids(self, session: ImportSession) -> list[int]:
        if session.period_start is None or session.period_end is None:
            if session.duplicate_sources:
                raise ValueError(
                    "Эти файлы уже использовались, но период нового отчета "
                    "не удалось определить. Импорт прерван."
                )
            return []

        run_ids = self.db.find_runs_by_period(session.period_start, session.period_end)
        duplicate_ids = {
            run_id
            for source in session.duplicate_sources
            for run_id in source.duplicate_run_ids
        }
        unrelated_duplicates = duplicate_ids.difference(run_ids)
        if unrelated_duplicates:
            raise ValueError(
                "Один из выбранных файлов уже использовался в отчете за другой период. "
                "Импорт прерван, чтобы не удваивать суммы."
            )
        return run_ids

    def complete_import(
        self,
        session: ImportSession,
        created_products: list[Product] | None = None,
        skipped_articles: set[str] | None = None,
        replace_run_ids: list[int] | None = None,
        source_period_warnings: list[str] | None = None,
        double_count_warnings: list[str] | None = None,
    ) -> RunCalculation:
        return self.complete_import_batch(
            [session],
            created_products=created_products,
            skipped_articles=skipped_articles,
            replace_run_ids_by_session=[list(replace_run_ids or [])],
            source_period_warnings_by_session=[list(source_period_warnings or [])],
            double_count_warnings_by_session=[list(double_count_warnings or [])],
        )[0]

    def complete_import_batch(
        self,
        sessions: list[ImportSession],
        created_products: list[Product] | None = None,
        skipped_articles: set[str] | None = None,
        replace_run_ids_by_session: list[list[int]] | None = None,
        source_period_warnings_by_session: list[list[str]] | None = None,
        double_count_warnings_by_session: list[list[str]] | None = None,
    ) -> list[RunCalculation]:
        if not sessions:
            raise ValueError("Нет подготовленных отчетов для сохранения")
        replacements = replace_run_ids_by_session or [[] for _session in sessions]
        warnings = source_period_warnings_by_session or [[] for _session in sessions]
        double_counts = double_count_warnings_by_session or [[] for _session in sessions]
        if (
            len(replacements) != len(sessions)
            or len(warnings) != len(sessions)
            or len(double_counts) != len(sessions)
        ):
            raise ValueError("Нарушена структура пакетного импорта")

        tax_rates = self.db.tax_rate_schedule()
        # Archived products remain valid calculation targets when their article
        # is present in an imported source row.
        product_map = self.db.product_map(active_only=False)
        for product in created_products or []:
            product_map[product.article] = product

        # Calculate every period before changing history. A malformed later month
        # therefore cannot leave a normally failed batch half-created.
        calculations: list[RunCalculation] = []
        for session, session_warnings, session_double_counts in zip(
            sessions, warnings, double_counts
        ):
            calculation = calculate_run(
                session.sources,
                product_map,
                tax_rates,
                skipped_articles=skipped_articles,
            )
            calculation.source_period_warnings = list(session_warnings)
            calculation.double_count_warnings = list(session_double_counts)
            calculations.append(calculation)

        for product in created_products or []:
            self.db.save_product(product, source="Новый артикул из отчета")
        for calculation, replace_run_ids in zip(calculations, replacements):
            stored_paths = self._store_source_files(calculation.source_files)
            calculation.run_id = self.db.save_run(
                calculation,
                stored_paths,
                replace_run_ids=replace_run_ids,
            )
        return calculations

    def _store_source_files(self, sources: list[ParsedSource]) -> dict[str, Path]:
        result: dict[str, Path] = {}
        for source in sources:
            safe_name = re.sub(r"[^\w.()\- ]+", "_", source.path.name, flags=re.UNICODE).strip()
            destination = self.paths["files"] / f"{source.file_hash[:12]}_{safe_name}"
            if not destination.exists():
                shutil.copy2(source.path, destination)
            result[str(source.path)] = destination
        return result

    def latest_run_id(self) -> int | None:
        runs = self.db.list_runs()
        return runs[-1].id if runs else None

    def recalculate_history(self) -> HistoryRecalculationResult:
        """Rebuild saved runs from stored source files while preserving historical inputs.

        Каждый отчет пересчитывается по своему сохраненному снимку налоговых ставок,
        поэтому налог меняется только через изменение справочника ставок с предпросмотром.
        """
        runs = self.db.list_runs()
        if not runs:
            return HistoryRecalculationResult(0, 0, 0.0, 0)

        current_products = self.db.product_map(active_only=False)
        replacements: list[RunReplacement] = []
        recovered_rows = 0
        financial_delta = 0.0
        units_delta = 0.0
        net_profit_delta = 0.0
        skipped_count = 0

        # Parse and calculate every report before replacing a single history row.
        # A missing or damaged source therefore leaves the whole history untouched.
        with tempfile.TemporaryDirectory(prefix="ozprice_history_recalc_") as temp_name:
            temp_root = Path(temp_name)
            for run in runs:
                old, calculation, stored_paths = self._rebuild_run(
                    run, temp_root, current_products
                )
                old_by_article = {item.article: item for item in old.products}
                for item in calculation.products:
                    previous = old_by_article.get(item.article)
                    if _result_has_activity(item) and (
                        previous is None or not _result_has_activity(previous)
                    ):
                        recovered_rows += 1
                new_totals = calculation.totals()
                old_totals = old.totals()
                financial_delta += new_totals["financial_result"] - old_totals["financial_result"]
                units_delta += new_totals["units"] - old_totals["units"]
                net_profit_delta += new_totals["net_profit"] - old_totals["net_profit"]
                skipped_count += len(calculation.skipped_articles)
                replacements.append(
                    RunReplacement(
                        old_run_id=run.id,
                        calculation=calculation,
                        stored_paths=stored_paths,
                        created_at=run.created_at,
                        planned_prices=self.db.planned_prices(run.id),
                    )
                )
            old_to_new = self.db.replace_runs(replacements)

        return HistoryRecalculationResult(
            replaced_runs=len(replacements),
            recovered_product_rows=recovered_rows,
            financial_result_delta=financial_delta,
            skipped_articles=skipped_count,
            units_delta=units_delta,
            net_profit_delta=net_profit_delta,
            old_to_new=old_to_new,
        )

    def preview_tax_rate_change(self, new_schedule: TaxRateSchedule) -> TaxRateChangePlan:
        """Найти отчеты, которых касается новый справочник ставок, и рассчитать их заново.

        Ничего не записывает: ни справочник, ни отчеты, ни файлы.
        """
        old_schedule = self.db.tax_rate_schedule()
        runs = self.db.list_runs()
        tax_periods = self.db.run_tax_periods()
        candidates = [
            run
            for run in runs
            if not self.db.run_tax_schedule(run.id).same_rates_between(
                new_schedule, *tax_periods.get(run.id, (None, None))
            )
        ]
        plan = TaxRateChangePlan(old_schedule=old_schedule, new_schedule=new_schedule)
        if not candidates:
            return plan

        current_products = self.db.product_map(active_only=False)
        with tempfile.TemporaryDirectory(prefix="ozprice_tax_preview_") as temp_name:
            temp_root = Path(temp_name)
            for run in candidates:
                old, calculation, stored_paths = self._rebuild_run(
                    run, temp_root, current_products, tax_rates=new_schedule
                )
                old_totals = old.totals()
                new_totals = calculation.totals()
                plan.items.append(
                    TaxRateChangeItem(
                        run_id=run.id,
                        report_name=run.report_name,
                        period_start=old.period_start,
                        period_end=old.period_end,
                        old_tax=old_totals["tax"],
                        new_tax=new_totals["tax"],
                        old_product_net_profit=old_totals["net_profit"],
                        new_product_net_profit=new_totals["net_profit"],
                        old_activity_net_profit=old_totals["report_net_profit"],
                        new_activity_net_profit=new_totals["report_net_profit"],
                    )
                )
                plan.replacements.append(
                    RunReplacement(
                        old_run_id=run.id,
                        calculation=calculation,
                        stored_paths=stored_paths,
                        created_at=run.created_at,
                        planned_prices=self.db.planned_prices(run.id),
                    )
                )
        return plan

    def apply_tax_rate_change(
        self,
        plan: TaxRateChangePlan,
        backup_path: Path | None = None,
    ) -> dict[int, int]:
        """Сохранить новый справочник и пересчитанные отчеты одной транзакцией.

        Перед пересчетом создается резервная копия, если передан ``backup_path``.
        """
        if self.db.tax_rate_schedule() != plan.old_schedule:
            raise ValueError(
                "Справочник ставок изменился после предпросмотра. Повторите изменение."
            )
        if plan.replacements and backup_path is not None:
            create_backup(self.paths["root"], backup_path)
        return self.db.replace_runs(plan.replacements, tax_schedule=plan.new_schedule)

    def _rebuild_run(
        self,
        run: RunSummary,
        temp_root: Path,
        current_products: dict[str, Product],
        tax_rates: TaxRateSchedule | None = None,
    ) -> tuple[RunCalculation, RunCalculation, dict[str, Path]]:
        """Пересчитать сохраненный отчет из копий исходных файлов с исторической себестоимостью.

        Без ``tax_rates`` используется снимок ставок самого отчета. Возвращает прежний и новый
        расчеты и соответствие исходных файлов их уже сохраненным копиям.
        """
        old = self.db.load_calculation(run.id)
        source_records = self.db.list_source_files(run.id)
        if not source_records:
            raise ValueError(
                f"У отчета «{run.report_name}» нет сохраненных исходных файлов. "
                "История не изменена."
            )

        run_root = temp_root / str(run.id)
        run_root.mkdir(parents=True)
        parsed_sources: list[ParsedSource] = []
        stored_paths: dict[str, Path] = {}
        for index, record in enumerate(source_records, start=1):
            stored_path = Path(str(record["stored_path"]))
            if not stored_path.is_file():
                raise ValueError(
                    f"Не найден сохраненный исходный файл «{record['original_name']}» "
                    f"для отчета «{run.report_name}». История не изменена."
                )
            original_name = Path(str(record["original_name"])).name
            destination = run_root / original_name
            if destination.exists():
                destination = run_root / f"{index}_{original_name}"
            shutil.copy2(stored_path, destination)
            source = parse_report(destination)
            source.duplicate_run_ids = self.db.find_runs_by_hash(source.file_hash)
            parsed_sources.append(source)
            stored_paths[str(source.path)] = stored_path

        historical_products = {
            item.article: Product(
                article=item.article,
                name=item.name,
                material_cost=item.material_cost,
                labor_cost=item.labor_cost,
                active=True,
                category=item.category,
            )
            for item in old.products
        }
        referenced_articles = {
            row.article
            for source in parsed_sources
            for row in source.accrual_rows
            if row.article
        } | {
            row.raw_article
            for source in parsed_sources
            for row in source.realization_rows
            if row.raw_article
        }
        for article in referenced_articles:
            if article not in historical_products and article in current_products:
                historical_products[article] = current_products[article]

        calculation = calculate_run(
            parsed_sources,
            historical_products,
            tax_rates or old.tax_schedule or old.tax_rate,
        )
        calculation.source_period_warnings = list(old.source_period_warnings)
        calculation.double_count_warnings = list(old.double_count_warnings)
        return old, calculation, stored_paths


def _result_has_activity(item) -> bool:
    return any(
        abs(value) > 0.000001
        for value in (
            item.units,
            item.revenue_no_points,
            item.partner_programs,
            item.points,
            item.commission,
            item.processing,
            item.delivery,
            item.logistics,
            item.reverse_logistics,
            item.returns_cancels,
            item.acquiring,
            item.stars,
            item.packaging,
            item.compensation,
            item.other,
            item.financial_result,
        )
    )


def _month_list(months: set[tuple[int, int]]) -> str:
    return ", ".join(f"{month:02d}.{year}" for year, month in sorted(months))


def split_import_sources(sources: list[ParsedSource]) -> list[ImportSession]:
    """Create one session per accrual file and match supplementary files by month."""
    accrual_sources = [source for source in sources if source.report_type == REPORT_ACCRUAL]
    realization_sources = [
        source for source in sources if source.report_type == REPORT_REALIZATION
    ]
    additional_income_sources = [
        source for source in sources if source.report_type == REPORT_ADDITIONAL_INCOME
    ]
    if not accrual_sources:
        raise ValueError("Для расчета нужен хотя бы один отчет по начислениям")

    sessions = [ImportSession(sources=[source], unknown_products=[]) for source in accrual_sources]
    sessions.sort(key=_session_sort_key)
    accruals_by_month: dict[tuple[int, int], list[ImportSession]] = {}
    for session in sessions:
        month = _session_month(session)
        if month is not None:
            accruals_by_month.setdefault(month, []).append(session)

    for source in sorted(realization_sources, key=_source_sort_key):
        month = _source_month(source)
        matches = accruals_by_month.get(month, []) if month is not None else []
        if len(matches) == 1:
            matches[0].sources.append(source)
            continue

        if len(sessions) == 1:
            # Preserve the explicit confirmation flow for a single accrual report.
            # ImportSession.realization_period_warnings() will record the mismatch.
            sessions[0].sources.append(source)
            continue

        period = _source_period(source)
        if month is None:
            raise ValueError(
                f"Не удалось определить один календарный месяц отчета по выкупам "
                f"«{source.path.name}» ({period}). При пакетном импорте невозможно "
                "надежно выбрать отчет по начислениям. Добавьте файл с однозначным "
                "периодом или импортируйте этот отчет отдельно."
            )
        if not matches:
            raise ValueError(
                f"Для отчета по выкупам «{source.path.name}» ({period}) не найден "
                f"отчет по начислениям за {month[1]:02d}.{month[0]}. "
                "Добавьте соответствующий отчет по начислениям или исключите этот "
                "файл из пакетного импорта."
            )
        names = ", ".join(f"«{session.sources[0].path.name}»" for session in matches)
        raise ValueError(
            f"Отчет по выкупам «{source.path.name}» ({period}) нельзя распределить "
            f"однозначно: за {month[1]:02d}.{month[0]} выбрано несколько отчетов "
            f"по начислениям: {names}. Импортируйте их раздельно."
        )

    for source in sorted(additional_income_sources, key=_source_sort_key):
        month = _source_month(source)
        matches = accruals_by_month.get(month, []) if month is not None else []
        if len(matches) == 1:
            matches[0].sources.append(source)
            continue
        period = _source_period(source)
        if month is None:
            raise ValueError(
                f"В акте дополнительного дохода «{source.path.name}» не удалось "
                f"определить календарный месяц ({period})."
            )
        if not matches:
            raise ValueError(
                f"Для акта дополнительного дохода «{source.path.name}» ({period}) "
                f"не найден отчет по начислениям за {month[1]:02d}.{month[0]}."
            )
        names = ", ".join(f"«{session.sources[0].path.name}»" for session in matches)
        raise ValueError(
            f"Акт дополнительного дохода «{source.path.name}» ({period}) нельзя "
            f"распределить однозначно: выбрано несколько отчетов за месяц: {names}."
        )

    return sessions


def _session_month(session: ImportSession) -> tuple[int, int] | None:
    if session.period_start is None or session.period_end is None:
        return None
    start = session.period_start
    end = session.period_end
    if (start.year, start.month) != (end.year, end.month):
        return None
    return start.year, start.month


def _source_month(source: ParsedSource) -> tuple[int, int] | None:
    if source.period_start is None or source.period_end is None:
        return None
    start = source.period_start
    end = source.period_end
    if (start.year, start.month) != (end.year, end.month):
        return None
    return start.year, start.month


def _source_period(source: ParsedSource) -> str:
    if source.period_start is None or source.period_end is None:
        return "период не определен"
    return f"{source.period_start:%d.%m.%Y}–{source.period_end:%d.%m.%Y}"


def _session_sort_key(session: ImportSession) -> tuple[date, date, str]:
    return (
        session.period_start or date.max,
        session.period_end or date.max,
        session.sources[0].path.name.casefold(),
    )


def _source_sort_key(source: ParsedSource) -> tuple[date, date, str]:
    return (
        source.period_start or date.max,
        source.period_end or date.max,
        source.path.name.casefold(),
    )
