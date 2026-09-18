from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from .backup import create_backup, inspect_backup, restore_backup, suggested_backup_name
from .aggregation import aggregate_calculations
from .calculator import calculate_scenario, discover_unknown_products
from .comparison import ComparisonMetric, compare_calculations
from .costs import (
    CostChange,
    CostEditorEntry,
    build_cost_changes,
    build_products_from_editor_entries,
    export_cost_catalog,
    read_cost_catalog,
)
from .config import APP_TITLE, APP_VERSION, save_storage_location
from .database import Database
from .excel_reader import (
    REPORT_ADDITIONAL_INCOME,
    REPORT_REALIZATION,
    preview_pdf,
    preview_sheet,
    workbook_sheet_names,
)
from .exporter import export_calculation, export_run, suggested_export_name
from .help_content import OVERVIEW_HELP_CONTENT, REPORTS_HELP_CONTENT, HelpContent
from .models import Product, ProductResult, RunCalculation, RunSummary, ScenarioRow, UnknownProduct
from .ordering import insert_at_group_end
from .service import AppService, ImportBatch, ImportSession
from .storage import migrate_storage
from .theme import apply_theme
from .trends import TrendPoint, build_trend_points, chart_bounds


THEME_LABELS = {"Системная": "system", "Темная": "dark", "Светлая": "light"}
THEME_VALUES = {value: key for key, value in THEME_LABELS.items()}
TREND_METRICS = {
    "Выручка": "revenue",
    "Чистая прибыль": "net_profit",
    "Доходность": "profitability",
    "Продажи, шт.": "units",
    "Нераспределенные доходы / расходы": "unallocated",
    "Средняя комиссия, % от выручки": "commission_share",
    "Логистика, % от выручки": "logistics_share",
    "Баллы, % от выручки": "points_share",
    "Чистая прибыль, % от выручки": "net_margin",
}
TREND_PERCENT_METRICS = {
    "profitability",
    "commission_share",
    "logistics_share",
    "points_share",
    "net_margin",
}
TREND_PERIOD_ALL = "Все годы"
TREND_PERIOD_SELECT = "Выбрать годы…"
CATEGORY_ALL = "Все категории"
CATEGORY_EMPTY = "Без категории"
SORT_NONE = "Без сортировки"
SORT_METRICS = (
    SORT_NONE,
    "Доходность",
    "Чистая прибыль",
    "Выручка",
    "Количество продаж",
)
SORT_ASCENDING = "По возрастанию (А-Я)"
SORT_DESCENDING = "По убыванию (Я-А)"


OVERVIEW_COLUMN_SPECS = (
    ("article", "Артикул", 120),
    ("name", "Наименование", 230),
    ("category", "Категория", 190),
    ("unit_cost", "Итого с/с", 125),
    ("material", "Материал", 125),
    ("labor", "Трудозатраты", 125),
    ("material_sold", "Материал проданного", 125),
    ("labor_sold", "Трудозатраты проданного", 125),
    ("cost_sold", "С/с проданного", 125),
    ("profitability", "Доходность", 125),
    ("net_unit", "Чистая прибыль на ед.", 125),
    ("profit_unit", "Прибыль от продаж на ед.", 125),
    ("net_total", "Чистая прибыль всего", 125),
    ("profit_total", "Прибыль от продаж всего", 125),
    ("avg_price", "Средняя цена", 125),
    ("tax", "Налог", 125),
    ("taxable", "Налогооблагаемый доход", 125),
    ("units", "Продажи", 125),
    ("revenue", "Выручка с баллами", 125),
    ("revenue_no_points", "Выручка без баллов", 125),
    ("partner", "Программы партнеров", 125),
    ("points", "Баллы", 125),
    ("commission", "Комиссия Ozon", 125),
    ("processing", "Обработка отправления", 125),
    ("delivery", "Доставка до ПВЗ", 125),
    ("logistics", "Логистика", 125),
    ("reverse", "Обратная логистика", 125),
    ("returns", "Возвраты/отмены", 125),
    ("acquiring", "Эквайринг", 125),
    ("stars", "Звездные товары", 125),
    ("packaging", "Упаковка и материалы", 125),
    ("compensation", "Компенсации Ozon", 125),
    ("other", "Прочие начисления", 125),
    ("financial_result", "Финрезультат Ozon", 125),
    ("commission_share", "Средняя комиссия, % от выручки", 205),
    ("logistics_share", "Логистика, % от выручки", 190),
    ("points_share", "Баллы, % от выручки", 175),
    ("net_margin", "Чистая прибыль, % от выручки", 215),
)

SCENARIO_COLUMN_SPECS = (
    ("article", "Артикул", 120),
    ("name", "Наименование", 230),
    ("category", "Категория", 190),
    ("cost", "Себестоимость", 135),
    ("units", "Продажи", 135),
    ("current_price", "Текущая цена", 135),
    ("planned_price", "Плановая цена", 135),
    ("change", "Изменение", 135),
    ("profitability", "Доходность", 135),
    ("other_costs", "Затраты Ozon без комиссии", 135),
    ("planned_revenue", "Плановая выручка", 135),
    ("commission_rate", "Средняя комиссия", 135),
    ("commission", "Плановая комиссия", 135),
    ("points", "Плановые баллы", 135),
    ("taxable", "Налоговая база", 135),
    ("tax", "Налог", 135),
    ("profit", "Прибыль от продаж", 135),
    ("profit_unit", "Прибыль/ед. до с/с", 135),
    ("net_unit", "Чистая прибыль/ед.", 135),
    ("net_total", "Чистая прибыль всего", 135),
)


class OZPriceAnalyzerApp(tk.Tk):
    def __init__(self, service: AppService | None = None):
        super().__init__()
        self.service = service or AppService()
        self.db = self.service.db
        self.current_run_id: int | None = None
        self.current_calculation: RunCalculation | None = None
        self.overview_calculation: RunCalculation | None = None
        self.overview_run_ids: set[int] = set()
        self.overview_selection_explicit = False
        self.overview_runs: list[RunSummary] = []
        self.overview_file_count = 0
        self.run_display_to_id: dict[str, int] = {}
        self.run_number_by_id: dict[int, int] = {}
        self.history_number_by_id: dict[int, int] = {}
        self.history_year_filter: set[int] | None = None
        self.history_year_filter_var = tk.StringVar(value="Все годы")
        self.trend_year_filter: set[int] | None = None
        self.trend_period_mode = "all"
        self.trend_period_var = tk.StringVar(value=TREND_PERIOD_ALL)
        self.trend_has_history = False
        self.source_by_iid: dict[str, dict[str, object]] = {}
        self.preview_headers: list[str] = []
        self.preview_rows: list[list[str]] = []
        self.preview_path: str | None = None
        self.scenario_rows: dict[str, ScenarioRow] = {}
        self.trend_points: list[TrendPoint] = []
        self.trend_canvas_points: list[tuple[float, float, TrendPoint]] = []
        self.import_in_progress = False
        self.import_queue: queue.Queue[tuple[ImportBatch | None, Exception | None]] = queue.Queue()
        self.colors = apply_theme(self, self.db.get_setting("theme", "system"))

        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1540x920")
        self.minsize(1180, 720)
        # A Tk font family containing spaces must be grouped as one Tcl list item.
        self.option_add("*Font", "{Segoe UI} 10")
        self._build_ui()
        self.refresh_all()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self._build_header()
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 12))

        self.overview_tab = ttk.Frame(self.notebook, padding=4)
        self.sources_tab = ttk.Frame(self.notebook, padding=4)
        self.breakdown_tab = ttk.Frame(self.notebook, padding=4)
        self.guide_tab = ttk.Frame(self.notebook, padding=4)
        self.scenario_tab = ttk.Frame(self.notebook, padding=4)
        self.history_tab = ttk.Frame(self.notebook, padding=4)
        self.trend_tab = ttk.Frame(self.notebook, padding=4)
        self.comparison_tab = ttk.Frame(self.notebook, padding=4)
        self.help_tab = ttk.Frame(self.notebook, padding=4)
        self.settings_tab = ttk.Frame(self.notebook, padding=4)
        self.notebook.add(self.overview_tab, text="Обзор")
        self.notebook.add(self.sources_tab, text="Исходные файлы")
        self.notebook.add(self.breakdown_tab, text="Разбивка")
        self.notebook.add(self.guide_tab, text="Справочник начислений")
        self.notebook.add(self.scenario_tab, text="Сценарий цены")
        self.notebook.add(self.history_tab, text="История отчетов")
        self.notebook.add(self.trend_tab, text="Динамика")
        self.notebook.add(self.comparison_tab, text="Сравнение периодов")
        self.notebook.add(self.help_tab, text="Справка")
        self.notebook.add(self.settings_tab, text="Настройки")

        self._build_overview_tab()
        self._build_sources_tab()
        self._build_breakdown_tab()
        self._build_guide_tab()
        self._build_scenario_tab()
        self._build_history_tab()
        self._build_trend_tab()
        self._build_comparison_tab()
        self._build_help_tab()
        self._build_settings_tab()

        self.status_var = tk.StringVar(value="Готово")
        ttk.Label(self, textvariable=self.status_var, style="Muted.TLabel").grid(
            row=2, column=0, sticky="ew", padx=22, pady=(0, 10)
        )

    def _build_header(self) -> None:
        header = ttk.Frame(self, padding=(22, 18, 22, 16))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        title_box = ttk.Frame(header)
        title_box.grid(row=0, column=0, sticky="w")
        ttk.Label(title_box, text="OZ Price Analyzer", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            title_box,
            text="Отчеты Ozon, история, контроль начислений и плановая доходность",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        actions = ttk.Frame(header)
        actions.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Label(actions, text="Отчет:", style="Muted.TLabel").grid(row=0, column=0, padx=(0, 6))
        self.run_var = tk.StringVar()
        self.run_combo = ttk.Combobox(actions, textvariable=self.run_var, state="readonly", width=32)
        self.run_combo.grid(row=0, column=1, padx=(0, 12))
        self.run_combo.bind("<<ComboboxSelected>>", self._on_run_selected)
        ttk.Button(actions, text="Импортировать отчеты", style="Accent.TButton", command=self.import_reports).grid(
            row=0, column=2, padx=5
        )
        ttk.Button(actions, text="Экспорт в Excel", command=self.export_current_run).grid(row=0, column=3, padx=5)
        ttk.Button(actions, text="О программе", command=self.show_about).grid(
            row=1, column=3, sticky="e", padx=5, pady=(6, 0)
        )

    def _build_overview_tab(self) -> None:
        self.overview_tab.columnconfigure(0, weight=1)
        self.overview_tab.rowconfigure(1, weight=1)
        self.overview_upper = ttk.Frame(self.overview_tab)
        self.overview_upper.grid(row=0, column=0, sticky="nsew")
        self.overview_upper.columnconfigure(0, weight=1)
        overview_header = ttk.Frame(self.overview_upper)
        overview_header.grid(row=0, column=0, sticky="ew", pady=(4, 4))
        overview_header.columnconfigure(1, weight=1)
        self.overview_header = overview_header
        ttk.Label(overview_header, text="Итоговый отчет", style="CompactSection.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.overview_scope_var = tk.StringVar(value="Текущий отчет")
        ttk.Label(
            overview_header,
            textvariable=self.overview_scope_var,
            style="Muted.TLabel",
        ).grid(row=0, column=1, sticky="e", padx=(20, 12))
        ttk.Button(
            overview_header,
            text="Выбрать отчеты…",
            style="CompactAccent.TButton",
            command=self.choose_overview_reports,
        ).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(
            overview_header,
            text="Только текущий",
            style="Compact.TButton",
            command=self.use_current_report_in_overview,
        ).grid(row=0, column=3)
        self.kpi_frame = ttk.Frame(self.overview_upper)
        self.kpi_frame.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        for column in range(6):
            self.kpi_frame.columnconfigure(column, weight=1)
        self.overview_totals_title_var = tk.StringVar(value="Итоги по отчету")
        ttk.Label(self.kpi_frame, textvariable=self.overview_totals_title_var, style="Muted.TLabel").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 2)
        )
        self.kpi_vars: dict[str, tk.StringVar] = {}
        cards = [
            ("revenue", "Выручка"),
            ("net_profit", "Чистая прибыль"),
            ("profitability", "Доходность"),
            ("units", "Продажи, шт."),
            ("unallocated", "Нераспределенные"),
            ("files", "Исходные файлы"),
        ]
        for index, (key, title) in enumerate(cards):
            self.kpi_vars[key] = tk.StringVar(value="—")
            card = ttk.Frame(self.kpi_frame, style="Card.TFrame", padding=(8, 5))
            card.grid(row=1, column=index, sticky="nsew", padx=(0 if index == 0 else 5, 0 if index == 5 else 5))
            ttk.Label(card, text=title, style="CompactCardMuted.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, textvariable=self.kpi_vars[key], style="CompactKpi.TLabel").grid(
                row=1, column=0, sticky="w", pady=(1, 0)
            )

        category_header = ttk.Frame(self.kpi_frame)
        category_header.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(4, 2))
        category_header.columnconfigure(0, weight=1)
        self.category_summary_title_var = tk.StringVar(value="Итоги по товарам выбранной категории")
        ttk.Label(
            category_header,
            textvariable=self.category_summary_title_var,
            style="CompactSection.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            category_header,
            text="Нераспределенные доходы / расходы сюда не включаются",
            style="Muted.TLabel",
        ).grid(row=0, column=1, sticky="e")

        self.category_kpi_vars: dict[str, tk.StringVar] = {}
        category_cards = [
            ("revenue", "Выручка"),
            ("net_profit", "Чистая прибыль"),
            ("profitability", "Доходность"),
            ("units", "Продажи, шт."),
            ("cost_sold", "С/с проданного"),
            ("financial_result", "Финрезультат Ozon"),
        ]
        for index, (key, title) in enumerate(category_cards):
            self.category_kpi_vars[key] = tk.StringVar(value="—")
            card = ttk.Frame(self.kpi_frame, style="Card.TFrame", padding=(8, 5))
            card.grid(
                row=3,
                column=index,
                sticky="nsew",
                padx=(0 if index == 0 else 5, 0 if index == 5 else 5),
            )
            ttk.Label(card, text=title, style="CompactCardMuted.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, textvariable=self.category_kpi_vars[key], style="CompactKpi.TLabel").grid(
                row=1, column=0, sticky="w", pady=(1, 0)
            )

        filters = ttk.Frame(self.overview_upper)
        filters.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        filters.columnconfigure(10, weight=1)
        ttk.Label(filters, text="Категория:").grid(row=0, column=0, padx=(0, 6))
        self.overview_category_var = tk.StringVar(value=CATEGORY_ALL)
        self.overview_category_combo = ttk.Combobox(
            filters,
            textvariable=self.overview_category_var,
            state="readonly",
            width=24,
            style="Compact.TCombobox",
        )
        self.overview_category_combo.grid(row=0, column=1, padx=(0, 14))
        self.overview_category_combo.bind("<<ComboboxSelected>>", lambda _event: self._populate_overview())
        ttk.Label(filters, text="Артикул:").grid(row=0, column=2, padx=(0, 6))
        self.overview_article_var = tk.StringVar()
        overview_search = ttk.Entry(
            filters,
            textvariable=self.overview_article_var,
            width=20,
            style="Compact.TEntry",
        )
        overview_search.grid(row=0, column=3, padx=(0, 14))
        overview_search.bind("<KeyRelease>", lambda _event: self._populate_overview())
        ttk.Label(filters, text="Сортировать:").grid(row=0, column=4, padx=(0, 6))
        self.overview_sort_var = tk.StringVar(value=SORT_NONE)
        overview_sort = ttk.Combobox(
            filters,
            textvariable=self.overview_sort_var,
            values=SORT_METRICS,
            state="readonly",
            width=21,
            style="Compact.TCombobox",
        )
        overview_sort.grid(row=0, column=5, padx=(0, 8))
        overview_sort.bind("<<ComboboxSelected>>", lambda _event: self._populate_overview())
        self.overview_sort_direction_var = tk.StringVar(value=SORT_ASCENDING)
        overview_direction = ttk.Combobox(
            filters,
            textvariable=self.overview_sort_direction_var,
            values=(SORT_ASCENDING, SORT_DESCENDING),
            state="readonly",
            width=24,
            style="Compact.TCombobox",
        )
        overview_direction.grid(row=0, column=6, padx=(0, 8))
        overview_direction.bind("<<ComboboxSelected>>", lambda _event: self._populate_overview())
        ttk.Button(
            filters,
            text="Сбросить",
            style="Compact.TButton",
            command=self._reset_overview_filters,
        ).grid(row=0, column=7)
        self.overview_count_var = tk.StringVar()
        ttk.Label(filters, textvariable=self.overview_count_var, style="Muted.TLabel").grid(
            row=0, column=10, sticky="e"
        )

        columns = [column for column, _heading, _width in OVERVIEW_COLUMN_SPECS]
        headings = [heading for _column, heading, _width in OVERVIEW_COLUMN_SPECS]
        widths = [width for _column, _heading, width in OVERVIEW_COLUMN_SPECS]
        self.overview_tree = self._create_tree(
            self.overview_tab, columns, headings, row=1, widths=widths
        )

    def _build_sources_tab(self) -> None:
        self.sources_tab.columnconfigure(0, weight=1)
        self.sources_tab.rowconfigure(3, weight=1)
        source_header = ttk.Frame(self.sources_tab)
        source_header.grid(row=0, column=0, sticky="ew", pady=(10, 8))
        source_header.columnconfigure(0, weight=1)
        ttk.Label(source_header, text="Файлы выбранного расчета", style="Section.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        source_columns = ["name", "type", "rows", "amount", "period", "hash"]
        source_headings = ["Файл", "Тип", "Строк", "Сумма", "Период", "SHA-256"]
        self.source_tree = self._create_tree(
            self.sources_tab, source_columns, source_headings, row=1, height=6, widths=[380, 150, 80, 130, 190, 220]
        )
        self.source_tree.bind("<<TreeviewSelect>>", self._on_source_selected)

        controls = ttk.Frame(self.sources_tab, padding=(0, 10, 0, 8))
        controls.grid(row=2, column=0, sticky="ew")
        controls.columnconfigure(5, weight=1)
        ttk.Button(
            controls,
            text="Просмотреть любой XLSX",
            style="Accent.TButton",
            command=self.browse_xlsx_preview,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=(0, 8))
        self.clear_preview_button = ttk.Button(
            controls,
            text="Очистить просмотр",
            command=self.clear_xlsx_preview,
            state="disabled",
        )
        self.clear_preview_button.grid(row=0, column=2, columnspan=2, sticky="w")
        ttk.Label(controls, text="Просмотр выполняется без запуска Excel", style="Muted.TLabel").grid(
            row=0, column=5, sticky="e", padx=(18, 0)
        )

        ttk.Label(controls, text="Лист:").grid(row=1, column=0, padx=(0, 6), pady=(10, 0))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(controls, textvariable=self.sheet_var, state="readonly", width=35)
        self.sheet_combo.grid(row=1, column=1, padx=(0, 18), pady=(10, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self._load_preview())
        ttk.Label(controls, text="Поиск в показанных строках:").grid(
            row=1, column=2, padx=(0, 6), pady=(10, 0)
        )
        self.preview_search_var = tk.StringVar()
        search = ttk.Entry(controls, textvariable=self.preview_search_var, width=35)
        search.grid(row=1, column=3, padx=(0, 8), pady=(10, 0))
        search.bind("<KeyRelease>", lambda _event: self._filter_preview())
        self.preview_file_var = tk.StringVar(value="Файл не выбран")
        ttk.Label(controls, textvariable=self.preview_file_var, style="Muted.TLabel").grid(
            row=2, column=0, columnspan=6, sticky="w", pady=(8, 0)
        )
        self.preview_container = ttk.Frame(self.sources_tab)
        self.preview_container.grid(row=3, column=0, sticky="nsew")
        self.preview_container.columnconfigure(0, weight=1)
        self.preview_container.rowconfigure(0, weight=1)
        self.preview_tree: ttk.Treeview | None = None

    def _build_breakdown_tab(self) -> None:
        self.breakdown_tab.columnconfigure(0, weight=1)
        self.breakdown_tab.rowconfigure(2, weight=1)
        ttk.Label(self.breakdown_tab, text="Нераспределенные доходы / расходы", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(10, 2)
        )
        ttk.Label(
            self.breakdown_tab,
            text="Здесь находятся начисления без артикула. Положительные суммы — доходы, отрицательные — расходы.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))
        self.breakdown_tree = self._create_tree(
            self.breakdown_tab,
            ["type", "count", "amount", "share"],
            ["Тип начисления", "Количество строк", "Сумма, руб.", "Доля в нераспределенных"],
            row=2,
            widths=[520, 150, 180, 190],
        )

    def _build_guide_tab(self) -> None:
        self.guide_tab.columnconfigure(0, weight=1)
        self.guide_tab.rowconfigure(2, weight=1)
        ttk.Label(self.guide_tab, text="Справочник начислений", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(10, 2)
        )
        ttk.Label(
            self.guide_tab,
            text="Фактическое наличие артикула проверяется для каждой строки. Справочник ничего не принуждает распределять.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))
        self.guide_tree = self._create_tree(
            self.guide_tab,
            ["type", "category", "current_status", "current_with", "current_without", "history_status", "history_with", "history_without"],
            ["Тип начисления", "Куда относится", "Текущий запуск", "С артикулом", "Без артикула", "История", "История с артикулом", "История без артикула"],
            row=2,
            widths=[320, 460, 260, 115, 115, 260, 150, 150],
        )

    def _build_scenario_tab(self) -> None:
        self.scenario_tab.columnconfigure(0, weight=1)
        self.scenario_tab.rowconfigure(4, weight=1)
        ttk.Label(self.scenario_tab, text="Доходность при плановой цене", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(10, 2)
        )
        ttk.Label(
            self.scenario_tab,
            text="Объем продаж остается текущим. Комиссия, баллы и прочие затраты пересчитываются по средним показателям выбранного периода.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))

        scenario_top = ttk.Frame(self.scenario_tab)
        scenario_top.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        scenario_top.columnconfigure(6, weight=1)
        ttk.Label(scenario_top, text="Плановая цена выбранного товара:").grid(row=0, column=0, padx=(0, 6))
        self.planned_price_var = tk.StringVar()
        ttk.Entry(scenario_top, textvariable=self.planned_price_var, width=16).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(scenario_top, text="Применить", command=self.apply_planned_price).grid(row=0, column=2, padx=(0, 18))
        ttk.Label(scenario_top, text="Изменить цену на, %:").grid(row=0, column=3, padx=(0, 6))
        self.batch_percent_var = tk.StringVar(value="5")
        ttk.Entry(scenario_top, textvariable=self.batch_percent_var, width=10).grid(row=0, column=4, padx=(0, 6))
        ttk.Button(
            scenario_top,
            text="Применить к выбранному",
            command=self.apply_selected_percent,
        ).grid(row=0, column=5, sticky="w", padx=(0, 6))
        ttk.Button(scenario_top, text="Применить ко всем", command=self.apply_batch_percent).grid(
            row=0, column=6, sticky="w"
        )
        ttk.Button(scenario_top, text="Сбросить цены", command=self.reset_scenario).grid(row=0, column=7, padx=(12, 0))

        filters = ttk.Frame(self.scenario_tab)
        filters.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        filters.columnconfigure(10, weight=1)
        ttk.Label(filters, text="Категория:").grid(row=0, column=0, padx=(0, 6))
        self.scenario_category_var = tk.StringVar(value=CATEGORY_ALL)
        self.scenario_category_combo = ttk.Combobox(
            filters, textvariable=self.scenario_category_var, state="readonly", width=24
        )
        self.scenario_category_combo.grid(row=0, column=1, padx=(0, 14))
        self.scenario_category_combo.bind("<<ComboboxSelected>>", lambda _event: self._populate_scenario())
        ttk.Label(filters, text="Артикул:").grid(row=0, column=2, padx=(0, 6))
        self.scenario_article_var = tk.StringVar()
        scenario_search = ttk.Entry(filters, textvariable=self.scenario_article_var, width=20)
        scenario_search.grid(row=0, column=3, padx=(0, 14))
        scenario_search.bind("<KeyRelease>", lambda _event: self._populate_scenario())
        ttk.Label(filters, text="Сортировать:").grid(row=0, column=4, padx=(0, 6))
        self.scenario_sort_var = tk.StringVar(value=SORT_NONE)
        scenario_sort = ttk.Combobox(
            filters, textvariable=self.scenario_sort_var, values=SORT_METRICS, state="readonly", width=21
        )
        scenario_sort.grid(row=0, column=5, padx=(0, 8))
        scenario_sort.bind("<<ComboboxSelected>>", lambda _event: self._populate_scenario())
        self.scenario_sort_direction_var = tk.StringVar(value=SORT_ASCENDING)
        scenario_direction = ttk.Combobox(
            filters,
            textvariable=self.scenario_sort_direction_var,
            values=(SORT_ASCENDING, SORT_DESCENDING),
            state="readonly",
            width=24,
        )
        scenario_direction.grid(row=0, column=6, padx=(0, 8))
        scenario_direction.bind("<<ComboboxSelected>>", lambda _event: self._populate_scenario())
        ttk.Button(filters, text="Сбросить", command=self._reset_scenario_filters).grid(row=0, column=7)
        self.scenario_count_var = tk.StringVar()
        ttk.Label(filters, textvariable=self.scenario_count_var, style="Muted.TLabel").grid(
            row=0, column=10, sticky="e"
        )

        self.scenario_tree = self._create_tree(
            self.scenario_tab,
            [column_id for column_id, _heading, _width in SCENARIO_COLUMN_SPECS],
            [heading for _column_id, heading, _width in SCENARIO_COLUMN_SPECS],
            row=4,
            widths=[width for _column_id, _heading, width in SCENARIO_COLUMN_SPECS],
        )
        self.scenario_tree.bind("<<TreeviewSelect>>", self._on_scenario_selected)

        self.scenario_kpi_frame = ttk.Frame(self.scenario_tab)
        self.scenario_kpi_frame.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        for column in range(4):
            self.scenario_kpi_frame.columnconfigure(column, weight=1)
        self.scenario_kpi_vars: dict[str, tk.StringVar] = {}
        for index, (key, title) in enumerate(
            [("current_revenue", "Текущая выручка"), ("planned_revenue", "Плановая выручка"), ("planned_net", "Плановая чистая прибыль"), ("planned_margin", "Плановая доходность")]
        ):
            self.scenario_kpi_vars[key] = tk.StringVar(value="—")
            card = ttk.Frame(self.scenario_kpi_frame, style="Card.TFrame", padding=(16, 12))
            card.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 5, 0 if index == 3 else 5))
            ttk.Label(card, text=title, style="CardMuted.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, textvariable=self.scenario_kpi_vars[key], style="Kpi.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))

    def _build_history_tab(self) -> None:
        self.history_tab.columnconfigure(0, weight=1)
        self.history_tab.rowconfigure(1, weight=0)
        self.history_tab.rowconfigure(3, weight=1)
        header = ttk.Frame(self.history_tab)
        header.grid(row=0, column=0, sticky="ew", pady=(10, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="История расчетов", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Показывать:", style="Muted.TLabel").grid(row=0, column=1, padx=(8, 4))
        ttk.Button(
            header,
            textvariable=self.history_year_filter_var,
            command=self.choose_history_years,
        ).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(header, text="Переименовать", command=self.rename_history_run).grid(row=0, column=3, padx=(8, 0))
        ttk.Button(header, text="Удалить", command=self.delete_history_run).grid(row=0, column=4, padx=(8, 0))
        self.history_tree = self._create_tree(
            self.history_tab,
            ["id", "name", "period", "created", "files", "units", "revenue", "net", "unallocated", "status"],
            ["№", "Наименование", "Период", "Дата расчета", "Файлов", "Продажи", "Выручка", "Чистая прибыль", "Нераспределенные", "Статус"],
            row=1,
            height=12,
            widths=[60, 300, 210, 160, 80, 110, 150, 150, 160, 100],
        )
        self.history_tree.bind("<<TreeviewSelect>>", self._on_history_selected)
        self.history_tree.bind("<Double-1>", self._open_history_run)
        ttk.Label(self.history_tab, text="Контроль качества выбранного отчета", style="Section.TLabel").grid(
            row=2, column=0, sticky="w", pady=(14, 8)
        )
        self.quality_tree = self._create_tree(
            self.history_tab,
            ["severity", "type", "message"],
            ["Уровень", "Проверка", "Сообщение"],
            row=3,
            widths=[140, 220, 900],
        )

    def _build_trend_tab(self) -> None:
        self.trend_tab.columnconfigure(0, weight=1)
        self.trend_tab.rowconfigure(3, weight=3)
        self.trend_tab.rowconfigure(5, weight=2)
        ttk.Label(self.trend_tab, text="Динамика показателей", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(10, 2)
        )
        ttk.Label(
            self.trend_tab,
            text="Каждая точка — сохраненный расчет. Периоды расположены по дате начала отчета.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))
        controls = ttk.Frame(self.trend_tab)
        controls.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(controls, text="Показатель:").grid(row=0, column=0, padx=(0, 6))
        self.trend_metric_var = tk.StringVar(value="Выручка")
        trend_combo = ttk.Combobox(
            controls,
            textvariable=self.trend_metric_var,
            state="readonly",
            values=list(TREND_METRICS),
            width=36,
        )
        trend_combo.grid(row=0, column=1, sticky="w")
        trend_combo.bind("<<ComboboxSelected>>", lambda _event: self._draw_trend_chart())
        ttk.Label(controls, text="Период:").grid(row=0, column=2, padx=(24, 6))
        self.trend_period_combo = ttk.Combobox(
            controls,
            textvariable=self.trend_period_var,
            state="readonly",
            values=(TREND_PERIOD_ALL, _current_year_period_label(), TREND_PERIOD_SELECT),
            width=24,
        )
        self.trend_period_combo.grid(row=0, column=3, sticky="w")
        self.trend_period_combo.bind("<<ComboboxSelected>>", self._on_trend_period_selected)

        self.trend_canvas = tk.Canvas(
            self.trend_tab,
            height=360,
            highlightthickness=1,
            bd=0,
        )
        self.trend_canvas.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        self.trend_canvas.bind("<Configure>", lambda _event: self._draw_trend_chart())
        self.trend_canvas.bind("<Motion>", self._trend_hover)
        self.trend_canvas.bind("<Leave>", lambda _event: self.trend_canvas.delete("tooltip"))

        ttk.Label(self.trend_tab, text="Таблица динамики", style="Section.TLabel").grid(
            row=4, column=0, sticky="w", pady=(4, 8)
        )
        self.trend_tree = self._create_tree(
            self.trend_tab,
            [
                "run", "period", "units", "revenue", "revenue_change", "net",
                "net_change", "profitability", "unallocated", "commission_share", "logistics_share",
                "points_share", "net_margin",
            ],
            [
                "№ отчета", "Период", "Продажи", "Выручка", "Изменение выручки",
                "Чистая прибыль", "Изменение прибыли", "Доходность", "Нераспределенные",
                "Средняя комиссия, % от выручки", "Логистика, % от выручки",
                "Баллы, % от выручки", "Чистая прибыль, % от выручки",
            ],
            row=5,
            widths=[80, 230, 110, 150, 170, 150, 170, 140, 170, 210, 190, 175, 215],
            height=8,
        )

    def _build_comparison_tab(self) -> None:
        self.comparison_tab.columnconfigure(0, weight=1)
        self.comparison_tab.rowconfigure(4, weight=1)
        ttk.Label(self.comparison_tab, text="Сравнение сохраненных периодов", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(10, 2)
        )
        ttk.Label(
            self.comparison_tab,
            text="Первый период — база сравнения. Изменение показывает второй период относительно первого.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))

        controls = ttk.Frame(self.comparison_tab)
        controls.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(controls, text="Первый период:").grid(row=0, column=0, padx=(0, 6))
        self.compare_first_var = tk.StringVar()
        self.compare_first_combo = ttk.Combobox(
            controls, textvariable=self.compare_first_var, state="readonly", width=34
        )
        self.compare_first_combo.grid(row=0, column=1, padx=(0, 18))
        ttk.Label(controls, text="Второй период:").grid(row=0, column=2, padx=(0, 6))
        self.compare_second_var = tk.StringVar()
        self.compare_second_combo = ttk.Combobox(
            controls, textvariable=self.compare_second_var, state="readonly", width=34
        )
        self.compare_second_combo.grid(row=0, column=3, padx=(0, 12))
        ttk.Button(controls, text="Сравнить", style="Accent.TButton", command=self.refresh_comparison).grid(
            row=0, column=4
        )

        self.comparison_kpi_frame = ttk.Frame(self.comparison_tab)
        self.comparison_kpi_frame.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        for column in range(4):
            self.comparison_kpi_frame.columnconfigure(column, weight=1)
        self.comparison_kpi_vars: dict[str, tk.StringVar] = {}
        for index, (key, title) in enumerate(
            [
                ("revenue", "Изменение выручки"),
                ("net_profit", "Изменение чистой прибыли"),
                ("units", "Изменение продаж"),
                ("unallocated", "Изменение нераспределенных"),
            ]
        ):
            self.comparison_kpi_vars[key] = tk.StringVar(value="—")
            card = ttk.Frame(self.comparison_kpi_frame, style="Card.TFrame", padding=(16, 12))
            card.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 5, 0 if index == 3 else 5))
            ttk.Label(card, text=title, style="CardMuted.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, textvariable=self.comparison_kpi_vars[key], style="Kpi.TLabel").grid(
                row=1, column=0, sticky="w", pady=(4, 0)
            )

        self.comparison_tree = self._create_tree(
            self.comparison_tab,
            [
                "article", "name", "units_first", "units_second", "units_change",
                "revenue_first", "revenue_second", "revenue_change", "revenue_percent",
                "profit_first", "profit_second", "profit_change", "profit_percent",
                "margin_first", "margin_second", "margin_change",
            ],
            [
                "Артикул", "Наименование", "Продажи 1", "Продажи 2", "Изменение продаж",
                "Выручка 1", "Выручка 2", "Изменение выручки", "Выручка, %",
                "Чистая прибыль 1", "Чистая прибыль 2", "Изменение прибыли", "Прибыль, %",
                "Доходность 1", "Доходность 2", "Изменение доходности",
            ],
            row=4,
            widths=[120, 230] + [135] * 14,
        )

    def _build_help_tab(self) -> None:
        self.help_tab.columnconfigure(0, weight=1)
        self.help_tab.rowconfigure(1, weight=1)
        ttk.Label(
            self.help_tab,
            text="Справка по исходным отчётам и расчётам",
            style="Section.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(10, 8))

        help_notebook = ttk.Notebook(self.help_tab)
        help_notebook.grid(row=1, column=0, sticky="nsew")
        reports_tab = ttk.Frame(help_notebook, padding=4)
        overview_tab = ttk.Frame(help_notebook, padding=4)
        help_notebook.add(reports_tab, text="Загрузка отчётов")
        help_notebook.add(overview_tab, text="Показатели «Обзора»")

        self.help_text_widgets = [
            self._create_help_document(reports_tab, REPORTS_HELP_CONTENT),
            self._create_help_document(overview_tab, OVERVIEW_HELP_CONTENT),
        ]
        self._configure_help_text_theme()

    def _create_help_document(self, parent: ttk.Frame, content: HelpContent) -> tk.Text:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        text = tk.Text(
            parent,
            wrap="word",
            relief="flat",
            borderwidth=0,
            padx=20,
            pady=16,
            cursor="arrow",
            takefocus=True,
        )
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        text.tag_configure(
            "title",
            font=("Segoe UI Semibold", 18),
            spacing1=2,
            spacing3=14,
        )
        text.tag_configure(
            "heading",
            font=("Segoe UI Semibold", 13),
            spacing1=16,
            spacing3=7,
        )
        text.tag_configure(
            "subheading",
            font=("Segoe UI Semibold", 11),
            spacing1=10,
            spacing3=4,
        )
        text.tag_configure("body", spacing1=2, spacing3=7, lmargin2=2)
        text.tag_configure("bullet", spacing1=2, spacing3=5, lmargin1=18, lmargin2=34)
        text.tag_configure(
            "formula",
            font=("Consolas", 10),
            spacing1=4,
            spacing3=7,
            lmargin1=18,
            lmargin2=18,
            rmargin=18,
        )
        text.tag_configure(
            "note",
            spacing1=5,
            spacing3=8,
            lmargin1=18,
            lmargin2=18,
            rmargin=18,
        )
        text.tag_configure(
            "warning",
            font=("Segoe UI Semibold", 10),
            spacing1=5,
            spacing3=8,
            lmargin1=18,
            lmargin2=18,
            rmargin=18,
        )
        for style, value in content:
            text.insert("end", f"{value}\n", style)
        text.configure(state="disabled")
        return text

    def _configure_help_text_theme(self) -> None:
        for text in getattr(self, "help_text_widgets", ()):
            text.configure(
                background=self.colors["surface"],
                foreground=self.colors["text"],
                insertbackground=self.colors["text"],
                selectbackground=self.colors["selection"],
                selectforeground=self.colors["text"],
            )
            text.tag_configure("title", foreground=self.colors["text"])
            text.tag_configure("heading", foreground=self.colors["accent"])
            text.tag_configure("subheading", foreground=self.colors["text"])
            text.tag_configure("body", foreground=self.colors["text"])
            text.tag_configure("bullet", foreground=self.colors["text"])
            text.tag_configure(
                "formula",
                foreground=self.colors["text"],
                background=self.colors["surface_alt"],
            )
            text.tag_configure("note", foreground=self.colors["muted"])
            text.tag_configure("warning", foreground=self.colors["warning"])

    def _build_settings_tab(self) -> None:
        self.settings_tab.columnconfigure(0, weight=1)
        self.settings_tab.rowconfigure(3, weight=1)
        ttk.Label(self.settings_tab, text="Настройки приложения", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(10, 8)
        )
        settings = ttk.Frame(self.settings_tab)
        settings.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(3, weight=1)

        ttk.Label(settings, text="Тема:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.theme_var = tk.StringVar(value=THEME_VALUES.get(self.db.get_setting("theme", "system"), "Системная"))
        theme_combo = ttk.Combobox(settings, textvariable=self.theme_var, state="readonly", values=list(THEME_LABELS), width=20)
        theme_combo.grid(row=0, column=1, sticky="w", pady=5)
        theme_combo.bind("<<ComboboxSelected>>", self._preview_theme)

        ttk.Label(settings, text="Налоговая ставка, %:").grid(row=0, column=2, sticky="w", padx=(24, 8), pady=5)
        self.tax_rate_var = tk.StringVar(value=_plain_number(float(self.db.get_setting("tax_rate", "0.04")) * 100))
        ttk.Entry(settings, textvariable=self.tax_rate_var, width=14).grid(row=0, column=3, sticky="w", pady=5)

        ttk.Label(settings, text="Повторный период:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Label(
            settings,
            text="Обновить существующий отчет или прервать импорт",
            style="Muted.TLabel",
        ).grid(row=1, column=1, sticky="w", pady=5)

        self.warn_realization_var = tk.BooleanVar(value=self.db.get_setting("warn_without_realization", "1") == "1")
        ttk.Checkbutton(
            settings,
            variable=self.warn_realization_var,
            text="Предупреждать, если не выбран отчет о выкупленных товарах",
        ).grid(row=1, column=2, columnspan=2, sticky="w", padx=(24, 0), pady=5)

        ttk.Label(settings, text="Строк в предпросмотре:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.preview_rows_var = tk.StringVar(value=self.db.get_setting("preview_rows", "500"))
        ttk.Spinbox(settings, textvariable=self.preview_rows_var, from_=100, to=5000, increment=100, width=12).grid(
            row=2, column=1, sticky="w", pady=5
        )
        ttk.Label(settings, text="Хранилище:").grid(row=2, column=2, sticky="w", padx=(24, 8), pady=5)
        storage_controls = ttk.Frame(settings)
        storage_controls.grid(row=2, column=3, sticky="ew", pady=5)
        storage_controls.columnconfigure(0, weight=1)
        self.storage_path_var = tk.StringVar(value=str(self.service.paths["root"]))
        ttk.Entry(storage_controls, textvariable=self.storage_path_var, state="readonly", width=48).grid(
            row=0, column=0, sticky="ew", padx=(0, 6)
        )
        ttk.Button(storage_controls, text="Изменить…", command=self.choose_storage_folder).grid(
            row=0, column=1, padx=3
        )
        ttk.Button(
            storage_controls,
            text="Открыть",
            command=lambda: _open_path(self.service.paths["root"]),
        ).grid(row=0, column=2, padx=(3, 0))
        ttk.Button(settings, text="Сохранить настройки", style="Accent.TButton", command=self.save_settings).grid(
            row=3, column=0, columnspan=4, sticky="w", pady=(10, 0)
        )

        backup_box = ttk.LabelFrame(settings, text="Резервная копия и перенос на другой компьютер", padding=(12, 10))
        backup_box.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(14, 0))
        backup_box.columnconfigure(0, weight=1)
        ttk.Label(
            backup_box,
            text="Архив содержит историю расчетов, настройки, себестоимость и сохраненные исходные отчеты.",
            style="Muted.TLabel",
        ).grid(row=0, column=0, sticky="w")
        backup_actions = ttk.Frame(backup_box)
        backup_actions.grid(row=0, column=1, sticky="e", padx=(16, 0))
        ttk.Button(
            backup_actions,
            text="Создать резервную копию",
            command=self.create_application_backup,
        ).grid(row=0, column=0, padx=4)
        ttk.Button(
            backup_actions,
            text="Восстановить / перенести",
            command=self.restore_application_backup,
        ).grid(row=0, column=1, padx=4)
        ttk.Button(
            backup_actions,
            text="Пересчитать историю",
            command=self.recalculate_saved_history,
        ).grid(row=0, column=2, padx=4)

        product_header = ttk.Frame(self.settings_tab)
        product_header.grid(row=2, column=0, sticky="ew", pady=(6, 8))
        product_header.columnconfigure(0, weight=1)
        ttk.Label(product_header, text="Товары и себестоимость", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(
            product_header,
            text="Редактировать справочник",
            style="Accent.TButton",
            command=self.open_cost_catalog_editor,
        ).grid(row=0, column=1, padx=(8, 4))
        ttk.Button(product_header, text="Добавить товар", command=self.add_product).grid(row=0, column=2, padx=4)
        ttk.Button(product_header, text="Изменить выбранный", command=self.edit_product).grid(row=0, column=3, padx=4)
        ttk.Button(product_header, text="В архив / восстановить", command=self.toggle_product).grid(row=0, column=4, padx=4)
        ttk.Button(product_header, text="Журнал изменений", command=self.show_cost_history).grid(row=0, column=5, padx=4)
        ttk.Label(
            product_header,
            text="При первом запуске справочник пуст. Загрузите XLSX, затем редактируйте товары прямо в приложении.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(4, 8))

        filters = ttk.Frame(product_header)
        filters.grid(row=2, column=0, columnspan=6, sticky="ew")
        filters.columnconfigure(5, weight=1)
        ttk.Label(filters, text="Поиск:").grid(row=0, column=0, padx=(0, 6))
        self.product_search_var = tk.StringVar()
        search_entry = ttk.Entry(filters, textvariable=self.product_search_var, width=28)
        search_entry.grid(row=0, column=1, padx=(0, 12))
        self.product_search_var.trace_add("write", lambda *_args: self.refresh_products())
        ttk.Label(filters, text="Показывать:").grid(row=0, column=2, padx=(0, 6))
        self.product_status_var = tk.StringVar(value="Все")
        status_combo = ttk.Combobox(
            filters,
            textvariable=self.product_status_var,
            state="readonly",
            values=("Все", "Активные", "Архив"),
            width=12,
        )
        status_combo.grid(row=0, column=3, padx=(0, 12))
        status_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_products())
        self.product_count_var = tk.StringVar(value="Показано: 0")
        ttk.Label(filters, textvariable=self.product_count_var, style="Muted.TLabel").grid(
            row=0, column=4, sticky="w"
        )
        ttk.Button(filters, text="Выгрузить XLSX", command=self.export_product_catalog).grid(row=0, column=6, padx=4)
        ttk.Button(filters, text="Загрузить XLSX", command=self.import_product_catalog).grid(row=0, column=7, padx=4)
        ttk.Button(filters, text="Очистить справочник", command=self.clear_product_catalog).grid(
            row=0, column=8, padx=(12, 4)
        )
        self.products_tree = self._create_tree(
            self.settings_tab,
            ["article", "name", "category", "total", "material", "labor", "status"],
            ["Артикул", "Наименование", "Категория", "Полная себестоимость", "Материал", "Трудозатраты", "Статус"],
            row=3,
            widths=[150, 320, 220, 180, 150, 150, 110],
        )
        self.products_tree.bind("<Double-1>", lambda _event: self.edit_product())

    def _create_tree(
        self,
        parent,
        columns: list[str],
        headings: list[str],
        row: int,
        widths: list[int] | None = None,
        height: int = 18,
    ) -> ttk.Treeview:
        container = ttk.Frame(parent)
        container.grid(row=row, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        tree = ttk.Treeview(container, columns=columns, show="headings", height=height)
        xscroll = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        yscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        widths = widths or [140] * len(columns)
        for column, heading, width in zip(columns, headings, widths):
            tree.heading(column, text=heading)
            tree.column(column, width=width, minwidth=70, stretch=False, anchor="w" if column in {"article", "name", "type", "category", "message"} else "e")
        return tree

    def refresh_all(self) -> None:
        self.refresh_products()
        self.refresh_runs()

    def refresh_runs(self) -> None:
        runs = self.db.list_runs()
        previous_runs = {run.id: run for run in self.overview_runs}
        missing_periods = {
            (previous_runs[run_id].period_start, previous_runs[run_id].period_end)
            for run_id in self.overview_run_ids
            if run_id in previous_runs and run_id not in {item.id for item in runs}
        }
        self.overview_runs = runs
        valid_ids = {run.id for run in runs}
        self.overview_run_ids.intersection_update(valid_ids)
        for period in missing_periods:
            replacements = [
                run.id for run in runs
                if (run.period_start, run.period_end) == period
            ]
            if replacements:
                self.overview_run_ids.add(replacements[-1])
        if self.overview_selection_explicit and not self.overview_run_ids:
            self.overview_selection_explicit = False
        self.run_display_to_id.clear()
        self.run_number_by_id = _run_positions(runs)
        values: list[str] = []
        for run in runs:
            period = _period_text(run.period_start, run.period_end)
            display = f"№{self.run_number_by_id[run.id]} · {run.report_name} · {period}"
            values.append(display)
            self.run_display_to_id[display] = run.id
        self.run_combo["values"] = values
        self.compare_first_combo["values"] = values
        self.compare_second_combo["values"] = values
        if not runs:
            self.run_var.set("Нет расчетов")
            self._clear_current_view()
        else:
            target_id = self.current_run_id if self.current_run_id in {run.id for run in runs} else runs[-1].id
            display = next(key for key, value in self.run_display_to_id.items() if value == target_id)
            self.run_var.set(display)
            self.select_run(target_id)
            if len(values) >= 2:
                if self.compare_first_var.get() not in values:
                    self.compare_first_var.set(values[-2])
                if self.compare_second_var.get() not in values:
                    self.compare_second_var.set(values[-1])
                self.refresh_comparison()
            else:
                self.compare_first_var.set(values[0])
                self.compare_second_var.set(values[0])
                self._clear_comparison("Для сравнения загрузите как минимум два периода")
        self.refresh_history()
        self.refresh_trends(runs)

    def select_run(self, run_id: int) -> None:
        self.current_run_id = run_id
        self.current_calculation = self.db.load_calculation(run_id)
        if not self.overview_selection_explicit:
            self.overview_run_ids = {run_id}
        self._refresh_active_report_views()
        self._populate_guide()
        self._populate_quality()
        status = f"Открыт отчет №{self._run_number(run_id)}: {_calculation_period(self.current_calculation)}"
        if self.overview_selection_explicit:
            status += f" · в обзоре отчетов: {len(self.overview_run_ids)}"
        self.status_var.set(status)

    def _on_run_selected(self, _event=None) -> None:
        run_id = self.run_display_to_id.get(self.run_var.get())
        if run_id is not None:
            # The most recently used selector owns the active report scope.
            # Choosing the top combobox therefore replaces an earlier multi-selection.
            self.overview_selection_explicit = False
            self.overview_run_ids = {run_id}
            self.select_run(run_id)

    def _active_report_run_ids(self) -> list[int]:
        """Return the active report scope in history order."""
        selected = set(self.__dict__.get("overview_run_ids", set()))
        ordered = [
            run.id
            for run in self.__dict__.get("overview_runs", [])
            if run.id in selected
        ]
        ordered.extend(sorted(selected.difference(ordered)))
        if not ordered and self.__dict__.get("current_run_id") is not None:
            ordered.append(self.current_run_id)
        return ordered

    def _active_report_calculation(self) -> RunCalculation | None:
        calculation = self.__dict__.get("overview_calculation")
        return calculation if calculation is not None else self.__dict__.get("current_calculation")

    def _active_single_run_id(self, *, notify: bool = False) -> int | None:
        run_ids = self._active_report_run_ids()
        if len(run_ids) == 1:
            return run_ids[0]
        if notify:
            messagebox.showinfo(
                "Сценарий цены",
                "Изменение плановых цен доступно при выборе одного отчета.\n\n"
                "Выберите отчет в верхнем списке или оставьте один отчет в окне "
                "«Выбрать отчеты…».",
                parent=self,
            )
        return None

    def _refresh_active_report_views(self) -> None:
        """Refresh every tab whose data follows the active report scope."""
        self._refresh_overview_calculation()
        self._populate_sources()
        self._populate_breakdown()
        self._populate_scenario()

    def _run_number(self, run_id: int | None) -> str:
        if run_id is None:
            return "—"
        number = self.run_number_by_id.get(run_id)
        return str(number) if number is not None else "—"

    def _current_run_status(self) -> str:
        if self.current_run_id is None:
            return "Готово"
        return f"Открыт отчет №{self._run_number(self.current_run_id)}"

    def _clear_current_view(self) -> None:
        self.current_run_id = None
        self.current_calculation = None
        self.overview_calculation = None
        self.overview_run_ids.clear()
        self.overview_selection_explicit = False
        self.overview_file_count = 0
        for tree in (self.overview_tree, self.source_tree, self.breakdown_tree, self.guide_tree, self.scenario_tree, self.quality_tree):
            tree.delete(*tree.get_children())
        self.clear_xlsx_preview()
        for variable in self.kpi_vars.values():
            variable.set("—")
        for variable in self.category_kpi_vars.values():
            variable.set("—")
        self.category_summary_title_var.set("Итоги по товарам выбранной категории")
        self.overview_scope_var.set("Нет выбранных отчетов")
        self.overview_totals_title_var.set("Итоги по отчету")
        for variable in self.scenario_kpi_vars.values():
            variable.set("—")
        self.overview_count_var.set("")
        self.scenario_count_var.set("")

    def _populate_overview(self) -> None:
        calculation = self.overview_calculation
        if calculation is None:
            return
        totals = calculation.totals()
        cost = totals["cost_sold"]
        self.kpi_vars["revenue"].set(_money(totals["revenue"]))
        self.kpi_vars["net_profit"].set(_money(totals["net_profit"]))
        self.kpi_vars["profitability"].set(_percent(totals["net_profit"] / cost if cost else 0))
        self.kpi_vars["units"].set(_number(totals["units"]))
        self.kpi_vars["unallocated"].set(_money(totals["unallocated"]))
        self.kpi_vars["files"].set(str(self.overview_file_count))
        _set_category_choices(
            self.overview_category_combo,
            self.overview_category_var,
            (result.category for result in calculation.products),
        )
        selected_category = self.overview_category_var.get()
        category_totals = summarize_category(
            calculation.products,
            calculation.tax_rate,
            selected_category,
        )
        category_name = "Все товары" if selected_category == CATEGORY_ALL else selected_category
        position_word = _russian_position_word(int(category_totals["product_count"]))
        self.category_summary_title_var.set(
            f"Итоги по категории: {category_name} · "
            f"{int(category_totals['product_count'])} {position_word}"
        )
        self.category_kpi_vars["revenue"].set(_money(category_totals["revenue"]))
        self.category_kpi_vars["net_profit"].set(_money(category_totals["net_profit"]))
        self.category_kpi_vars["profitability"].set(_percent(category_totals["profitability"]))
        self.category_kpi_vars["units"].set(_number(category_totals["units"]))
        self.category_kpi_vars["cost_sold"].set(_money(category_totals["cost_sold"]))
        self.category_kpi_vars["financial_result"].set(_money(category_totals["financial_result"]))
        visible = filter_product_results(
            calculation.products,
            calculation.tax_rate,
            category=self.overview_category_var.get(),
            article_query=self.overview_article_var.get(),
            sort_metric=self.overview_sort_var.get(),
            descending=self.overview_sort_direction_var.get() == SORT_DESCENDING,
        )
        self.overview_tree.delete(*self.overview_tree.get_children())
        for result in visible:
            values = _result_values(result, calculation.tax_rate)
            tag = "negative" if result.net_profit(calculation.tax_rate) < 0 else "positive"
            self.overview_tree.insert("", "end", iid=result.article, values=values, tags=(tag,))
        self.overview_count_var.set(f"Показано: {len(visible)} из {len(calculation.products)}")
        self._configure_value_tags(self.overview_tree)

    def choose_overview_reports(self) -> None:
        runs = self.db.list_runs()
        if not runs:
            messagebox.showinfo("Период обзора", "В истории пока нет отчетов", parent=self)
            return
        selected = self.overview_run_ids or ({self.current_run_id} if self.current_run_id else set())
        dialog = OverviewReportSelectionDialog(self, runs, selected)
        self.wait_window(dialog)
        if not dialog.confirmed:
            return
        self.overview_run_ids = set(dialog.selected_run_ids)
        self.overview_selection_explicit = True
        self._refresh_active_report_views()
        self.status_var.set(
            f"Обзор сформирован по {len(self.overview_run_ids)} отчетам: "
            f"{_calculation_period(self.overview_calculation)}"
        )

    def use_current_report_in_overview(self) -> None:
        if self.current_run_id is None:
            return
        self.overview_selection_explicit = False
        self.overview_run_ids = {self.current_run_id}
        self._refresh_active_report_views()
        self.status_var.set(
            f"Обзор показывает текущий отчет №{self._run_number(self.current_run_id)}"
        )

    def _refresh_overview_calculation(self) -> None:
        selected_runs = [
            run for run in self.overview_runs if run.id in self.overview_run_ids
        ]
        if not selected_runs and self.current_run_id is not None:
            selected_runs = [
                run for run in self.overview_runs if run.id == self.current_run_id
            ]
            self.overview_run_ids = {self.current_run_id}
            self.overview_selection_explicit = False
        if not selected_runs:
            self.overview_calculation = None
            self.overview_file_count = 0
            return

        if (
            len(selected_runs) == 1
            and self.current_calculation is not None
            and selected_runs[0].id == self.current_run_id
        ):
            self.overview_calculation = self.current_calculation
        else:
            self.overview_calculation = aggregate_calculations(
                [self.db.load_calculation(run.id) for run in selected_runs]
            )
        self.overview_file_count = sum(run.source_count for run in selected_runs)
        calculation = self.overview_calculation
        count = len(selected_runs)
        self.overview_scope_var.set(
            f"Период: {_calculation_period(calculation)} · {_russian_report_count(count)}"
        )
        self.overview_totals_title_var.set(
            "Итоги по выбранному периоду" if count > 1 else "Итоги по отчету"
        )
        self._populate_overview()

    def _reset_overview_filters(self) -> None:
        self.overview_category_var.set(CATEGORY_ALL)
        self.overview_article_var.set("")
        self.overview_sort_var.set(SORT_NONE)
        self.overview_sort_direction_var.set(SORT_ASCENDING)
        self._populate_overview()

    def _populate_sources(self) -> None:
        self.source_tree.delete(*self.source_tree.get_children())
        self.source_by_iid.clear()
        sources: list[tuple[int, dict[str, object]]] = []
        for run_id in self._active_report_run_ids():
            sources.extend((run_id, row) for row in self.db.list_source_files(run_id))
        for run_id, row in sources:
            iid = f"{run_id}:{row['id']}"
            self.source_by_iid[iid] = row
            report_type = {
                "ACCRUAL": "Начисления",
                "REALIZATION": "Выкупленные товары",
                REPORT_ADDITIONAL_INCOME: "Дополнительный доход",
            }.get(str(row["report_type"]), str(row["report_type"]))
            period = _period_text(row.get("period_start"), row.get("period_end"))
            self.source_tree.insert(
                "",
                "end",
                iid=iid,
                values=(row["original_name"], report_type, row["row_count"], _money(float(row["total_amount"])), period, str(row["file_hash"])[:24]),
            )
        if sources:
            first_run_id, first_row = sources[0]
            first = f"{first_run_id}:{first_row['id']}"
            self.source_tree.selection_set(first)
            self.source_tree.focus(first)
            self._on_source_selected()
        else:
            self.clear_xlsx_preview()

    def _on_source_selected(self, _event=None) -> None:
        selection = self.source_tree.selection()
        if not selection:
            return
        source = self.source_by_iid.get(selection[0])
        if source is None:
            return
        try:
            self.preview_path = str(source["stored_path"])
            self.preview_file_var.set(f"Файл: {source['original_name']}")
            if str(source["report_type"]) == REPORT_ADDITIONAL_INCOME:
                self.sheet_combo["values"] = ("PDF",)
                self.sheet_var.set("PDF")
                max_rows = int(self.db.get_setting("preview_rows", "500"))
                self.preview_headers, self.preview_rows = preview_pdf(
                    self.preview_path,
                    max_rows=max_rows,
                )
                self._filter_preview()
                self.clear_preview_button.configure(state="normal")
                return
            sheets = workbook_sheet_names(self.preview_path)
            self.sheet_combo["values"] = sheets
            preferred = str(source["sheet_name"])
            self.sheet_var.set(preferred if preferred in sheets else sheets[0])
            self._load_preview()
        except Exception as exc:
            messagebox.showerror("Просмотр файла", str(exc), parent=self)

    def _load_preview(self) -> None:
        if not self.preview_path or not self.sheet_var.get():
            return
        max_rows = int(self.db.get_setting("preview_rows", "500"))
        try:
            self.preview_headers, self.preview_rows = preview_sheet(
                self.preview_path, self.sheet_var.get(), max_rows=max_rows
            )
            self._filter_preview()
            self.clear_preview_button.configure(state="normal")
        except Exception as exc:
            messagebox.showerror("Просмотр файла", str(exc), parent=self)

    def browse_xlsx_preview(self) -> None:
        path = filedialog.askopenfilename(
            title="Просмотреть книгу без Excel",
            filetypes=[("Книги Excel", "*.xlsx")],
            parent=self,
        )
        if not path:
            return
        try:
            sheets = workbook_sheet_names(path)
            if not sheets:
                raise ValueError("В книге нет листов")
            self.preview_path = path
            self.preview_file_var.set(f"Файл: {Path(path).name} · только просмотр")
            self.sheet_combo["values"] = sheets
            self.sheet_var.set(sheets[0])
            self.source_tree.selection_remove(*self.source_tree.selection())
            self._load_preview()
        except Exception as exc:
            messagebox.showerror("Просмотр файла", str(exc), parent=self)

    def clear_xlsx_preview(self) -> None:
        self.preview_path = None
        self.preview_headers = []
        self.preview_rows = []
        self.preview_search_var.set("")
        self.sheet_var.set("")
        self.sheet_combo["values"] = ()
        self.preview_file_var.set("Файл не выбран")
        selection = self.source_tree.selection()
        if selection:
            self.source_tree.selection_remove(*selection)
        if self.preview_tree is not None:
            self.preview_tree.master.destroy()
            self.preview_tree = None
        self.clear_preview_button.configure(state="disabled")

    def _filter_preview(self) -> None:
        query = self.preview_search_var.get().casefold().strip()
        rows = self.preview_rows
        if query:
            rows = [row for row in rows if query in " | ".join(row).casefold()]
        self._render_preview_tree(self.preview_headers, rows)

    def _render_preview_tree(self, headers: list[str], rows: list[list[str]]) -> None:
        if self.preview_tree is not None:
            self.preview_tree.master.destroy()
        if not headers:
            self.preview_tree = None
            return
        columns = [f"c{index}" for index in range(len(headers))]
        frame = ttk.Frame(self.preview_container)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        for column, heading in zip(columns, headers):
            tree.heading(column, text=heading)
            tree.column(column, width=75 if heading == "Строка" else 150, minwidth=55, stretch=False, anchor="w")
        for row in rows:
            tree.insert("", "end", values=row)
        self.preview_tree = tree

    def _populate_breakdown(self) -> None:
        calculation = self._active_report_calculation()
        self.breakdown_tree.delete(*self.breakdown_tree.get_children())
        if calculation is None:
            return
        total = calculation.unallocated_total
        for accrual_type, (count, amount) in calculation.unallocated.items():
            share = amount / total if total else 0.0
            tag = "negative" if amount < 0 else "positive"
            self.breakdown_tree.insert("", "end", values=(accrual_type, count, _money(amount), _percent(share)), tags=(tag,))
        self.breakdown_tree.insert("", "end", values=("Итого", sum(x[0] for x in calculation.unallocated.values()), _money(total), _percent(1 if total else 0)), tags=("total",))
        if calculation.taxable_unallocated_income:
            self.breakdown_tree.insert(
                "",
                "end",
                values=(
                    "Налог с нераспределенных доходов",
                    "—",
                    _money(-calculation.unallocated_income_tax),
                    "—",
                ),
                tags=("negative",),
            )
            self.breakdown_tree.insert(
                "",
                "end",
                values=(
                    "Итого после налога с доходов",
                    "—",
                    _money(total - calculation.unallocated_income_tax),
                    "—",
                ),
                tags=("total",),
            )
        self._configure_value_tags(self.breakdown_tree)

    def _populate_guide(self) -> None:
        if self.current_run_id is None:
            return
        self.guide_tree.delete(*self.guide_tree.get_children())
        for item in self.db.accrual_guide(self.current_run_id):
            self.guide_tree.insert(
                "",
                "end",
                values=(
                    item["accrual_type"], item["category"], item["current_status"], item["current_with"],
                    item["current_without"], item["history_status"], item["history_with"], item["history_without"],
                ),
            )

    def _populate_scenario(self) -> None:
        calculation = self._active_report_calculation()
        self.scenario_tree.delete(*self.scenario_tree.get_children())
        self.scenario_rows.clear()
        if calculation is None:
            for variable in self.scenario_kpi_vars.values():
                variable.set("—")
            self.scenario_count_var.set("")
            return
        target_run_id = self._active_single_run_id()
        prices = self.db.planned_prices(target_run_id) if target_run_id is not None else {}
        planned_revenue_total = 0.0
        planned_net_total = 0.0
        planned_cost_total = 0.0
        scenarios: list[ScenarioRow] = []
        for result in calculation.products:
            scenario = calculate_scenario(result, calculation.tax_rate, prices.get(result.article))
            scenarios.append(scenario)
            self.scenario_rows[result.article] = scenario
            if scenario.planned_revenue is not None:
                planned_revenue_total += scenario.planned_revenue
            if scenario.net_profit_per_unit is not None:
                planned_net_total += scenario.net_profit_per_unit * scenario.units
                planned_cost_total += scenario.unit_cost * scenario.units
        _set_category_choices(
            self.scenario_category_combo,
            self.scenario_category_var,
            (row.category for row in scenarios),
        )
        visible = filter_scenario_rows(
            scenarios,
            category=self.scenario_category_var.get(),
            article_query=self.scenario_article_var.get(),
            sort_metric=self.scenario_sort_var.get(),
            descending=self.scenario_sort_direction_var.get() == SORT_DESCENDING,
        )
        for scenario in visible:
            tag = "negative" if (scenario.net_profit_per_unit or 0) < 0 else "positive"
            self.scenario_tree.insert(
                "", "end", iid=scenario.article, values=_scenario_values(scenario), tags=(tag,)
            )
        totals = calculation.totals()
        self.scenario_kpi_vars["current_revenue"].set(_money(totals["revenue"]))
        self.scenario_kpi_vars["planned_revenue"].set(_money(planned_revenue_total))
        self.scenario_kpi_vars["planned_net"].set(_money(planned_net_total))
        self.scenario_kpi_vars["planned_margin"].set(_percent(planned_net_total / planned_cost_total if planned_cost_total else 0))
        scope_note = ""
        active_count = len(self._active_report_run_ids())
        if active_count > 1:
            scope_note = f" · {_russian_report_count(active_count)} · цены только для 1 отчета"
        self.scenario_count_var.set(
            f"Показано: {len(visible)} из {len(scenarios)}{scope_note}"
        )
        self._configure_value_tags(self.scenario_tree)

    def _reset_scenario_filters(self) -> None:
        self.scenario_category_var.set(CATEGORY_ALL)
        self.scenario_article_var.set("")
        self.scenario_sort_var.set(SORT_NONE)
        self.scenario_sort_direction_var.set(SORT_ASCENDING)
        self._populate_scenario()

    def _on_scenario_selected(self, _event=None) -> None:
        selection = self.scenario_tree.selection()
        if not selection:
            return
        row = self.scenario_rows.get(selection[0])
        self.planned_price_var.set(_plain_number(row.planned_price) if row and row.planned_price is not None else "")

    def apply_planned_price(self) -> None:
        run_id = self._active_single_run_id(notify=True)
        if run_id is None:
            return
        selection = self.scenario_tree.selection()
        if not selection:
            messagebox.showinfo("Плановая цена", "Сначала выберите товар в таблице", parent=self)
            return
        try:
            value = _parse_number(self.planned_price_var.get())
            if value < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Плановая цена", "Введите неотрицательную цену", parent=self)
            return
        self.db.save_planned_price(run_id, selection[0], value)
        self._populate_scenario()
        self.scenario_tree.selection_set(selection[0])

    def apply_batch_percent(self) -> None:
        run_id = self._active_single_run_id(notify=True)
        calculation = self._active_report_calculation()
        if run_id is None or calculation is None:
            return
        percent = self._scenario_percent()
        if percent is None:
            return
        for result in calculation.products:
            current = result.average_price()
            if current is not None:
                self.db.save_planned_price(run_id, result.article, current * (1 + percent))
        self._populate_scenario()

    def apply_selected_percent(self) -> None:
        run_id = self._active_single_run_id(notify=True)
        if run_id is None:
            return
        selection = self.scenario_tree.selection()
        if not selection:
            messagebox.showinfo("Изменение цены", "Сначала выберите товар в таблице", parent=self)
            return
        article = selection[0]
        row = self.scenario_rows.get(article)
        if row is None or row.current_price is None:
            messagebox.showinfo(
                "Изменение цены",
                "Для выбранного товара нет текущей средней цены",
                parent=self,
            )
            return
        percent = self._scenario_percent()
        if percent is None:
            return
        self.db.save_planned_price(run_id, article, row.current_price * (1 + percent))
        self._populate_scenario()
        self.scenario_tree.selection_set(article)
        self.scenario_tree.focus(article)
        self.scenario_tree.see(article)

    def _scenario_percent(self) -> float | None:
        try:
            percent = _parse_number(self.batch_percent_var.get()) / 100
            if percent <= -1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Изменение цен", "Введите процент больше -100", parent=self)
            return None
        return percent

    def reset_scenario(self) -> None:
        run_id = self._active_single_run_id(notify=True)
        if run_id is None:
            return
        if messagebox.askyesno("Сбросить сценарий", "Вернуть плановые цены к текущим средним?", parent=self):
            self.db.clear_planned_prices(run_id)
            self._populate_scenario()

    def refresh_history(self, selected_run_id: int | None = None) -> None:
        if selected_run_id is None:
            selection = self.history_tree.selection()
            selected_run_id = int(selection[0]) if selection else None
        self.history_tree.delete(*self.history_tree.get_children())
        runs = self.db.list_runs()
        available_years = sorted({year for run in runs for year in _run_years(run)})
        if self.history_year_filter is not None:
            self.history_year_filter.intersection_update(available_years)
            if not self.history_year_filter:
                self.history_year_filter = None
        visible_runs = _filter_runs_by_years(runs, self.history_year_filter)
        self.history_number_by_id = _run_positions(visible_runs)
        self.history_year_filter_var.set(_year_filter_label(self.history_year_filter))
        for run in visible_runs:
            self.history_tree.insert(
                "",
                "end",
                iid=str(run.id),
                values=(
                    self.history_number_by_id[run.id], run.report_name,
                    _period_text(run.period_start, run.period_end), run.created_at[:16], run.source_count,
                    _number(run.units), _money(run.revenue), _money(run.net_profit), _money(run.unallocated_total), run.status,
                ),
            )
        visible_ids = {run.id for run in visible_runs}
        if selected_run_id not in visible_ids:
            selected_run_id = visible_runs[0].id if visible_runs else None
        selected_iid = str(selected_run_id) if selected_run_id is not None else ""
        if selected_iid and selected_iid in self.history_tree.get_children():
            self.history_tree.selection_set(selected_iid)
            self.history_tree.focus(selected_iid)
            self._populate_quality(selected_run_id)
        else:
            self._populate_quality(None, use_current=False)

    def _on_history_selected(self, _event=None) -> None:
        selection = self.history_tree.selection()
        self._populate_quality(int(selection[0]) if selection else None, use_current=False)

    def choose_history_years(self) -> None:
        runs = self.db.list_runs()
        years = sorted({year for run in runs for year in _run_years(run)})
        if not years:
            messagebox.showinfo("История отчетов", "В истории пока нет отчетов", parent=self)
            return
        dialog = HistoryYearFilterDialog(self, years, self.history_year_filter)
        self.wait_window(dialog)
        if dialog.cancelled:
            return
        self.history_year_filter = dialog.selected_years
        self.refresh_history()

    def rename_history_run(self) -> None:
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showinfo("История отчетов", "Выберите отчет для переименования", parent=self)
            return
        run_id = int(selection[0])
        current_name = str(self.history_tree.item(selection[0], "values")[1])
        new_name = simpledialog.askstring(
            "Переименовать отчет",
            "Новое наименование:",
            initialvalue=current_name,
            parent=self,
        )
        if new_name is None:
            return
        try:
            self.db.rename_run(run_id, new_name)
        except (ValueError, KeyError) as exc:
            messagebox.showerror("Переименовать отчет", str(exc), parent=self)
            return
        self.refresh_runs()
        self.refresh_history(run_id)
        number = self.history_number_by_id.get(run_id)
        self.status_var.set(f"Отчет №{number} переименован" if number else "Отчет переименован")

    def delete_history_run(self) -> None:
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showinfo("История отчетов", "Выберите отчет для удаления", parent=self)
            return
        run_id = int(selection[0])
        values = self.history_tree.item(selection[0], "values")
        report_name = str(values[1])
        period = str(values[2])
        if not messagebox.askyesno(
            "Удалить отчет",
            f"Удалить «{report_name}» ({period}) из истории?\n\n"
            "Расчет, его сценарные цены и контроль качества будут удалены без возможности восстановления.",
            icon="warning",
            parent=self,
        ):
            return
        try:
            removed_files = self.db.delete_run(run_id)
        except KeyError as exc:
            messagebox.showerror("Удалить отчет", str(exc), parent=self)
            return
        if self.current_run_id == run_id:
            self.current_run_id = None
            self.current_calculation = None
        self.refresh_runs()
        suffix = f"; удалено копий исходных файлов: {removed_files}" if removed_files else ""
        self.status_var.set(f"Отчет удален{suffix}")

    def refresh_trends(self, runs=None) -> None:
        all_runs = list(runs) if runs is not None else self.db.list_runs()
        self.trend_has_history = bool(all_runs)
        selected_years = (
            {date.today().year}
            if self.trend_period_mode == "current"
            else self.trend_year_filter
        )
        self.trend_period_var.set(_trend_period_label(self.trend_period_mode, selected_years))
        self.trend_points = build_trend_points(_filter_runs_by_years(all_runs, selected_years))
        self.trend_tree.delete(*self.trend_tree.get_children())
        previous: TrendPoint | None = None
        for point in self.trend_points:
            revenue_change = point.revenue - previous.revenue if previous else None
            profit_change = point.net_profit - previous.net_profit if previous else None
            tag = "positive" if profit_change is None or profit_change >= 0 else "negative"
            self.trend_tree.insert(
                "",
                "end",
                iid=str(point.run_id),
                values=(
                    self._run_number(point.run_id),
                    point.label,
                    _number(point.units),
                    _money(point.revenue),
                    _signed_money(revenue_change) if revenue_change is not None else "—",
                    _money(point.net_profit),
                    _signed_money(profit_change) if profit_change is not None else "—",
                    _percent(point.profitability),
                    _money(point.unallocated),
                    _percent(point.commission_share),
                    _percent(point.logistics_share),
                    _percent(point.points_share),
                    _percent(point.net_margin),
                ),
                tags=(tag,),
            )
            previous = point
        self._configure_value_tags(self.trend_tree)
        self._draw_trend_chart()

    def _on_trend_period_selected(self, _event=None) -> None:
        selection = self.trend_period_var.get()
        if selection == TREND_PERIOD_ALL:
            self.trend_period_mode = "all"
            self.trend_year_filter = None
            self.refresh_trends()
            return
        if selection == _current_year_period_label():
            self.trend_period_mode = "current"
            self.trend_year_filter = None
            self.refresh_trends()
            return
        if selection == TREND_PERIOD_SELECT:
            self.choose_trend_years()

    def choose_trend_years(self) -> None:
        runs = self.db.list_runs()
        years = sorted({year for run in runs for year in _run_years(run)})
        if not years:
            self.trend_period_var.set(_trend_period_label(self.trend_period_mode, self.trend_year_filter))
            messagebox.showinfo("Динамика", "В истории пока нет отчетов", parent=self)
            return
        selected_years = self.trend_year_filter if self.trend_period_mode == "custom" else None
        dialog = HistoryYearFilterDialog(
            self,
            years,
            selected_years,
            title="Период динамики по годам",
        )
        self.wait_window(dialog)
        if dialog.cancelled:
            active_years = (
                {date.today().year}
                if self.trend_period_mode == "current"
                else self.trend_year_filter
            )
            self.trend_period_var.set(_trend_period_label(self.trend_period_mode, active_years))
            return
        self.trend_year_filter = dialog.selected_years
        self.trend_period_mode = "all" if dialog.selected_years is None else "custom"
        self.refresh_trends(runs)

    def _draw_trend_chart(self) -> None:
        if not hasattr(self, "trend_canvas"):
            return
        canvas = self.trend_canvas
        canvas.delete("all")
        canvas.configure(background=self.colors["surface"], highlightbackground=self.colors["border"])
        self.trend_canvas_points.clear()
        width = max(canvas.winfo_width(), 680)
        height = max(canvas.winfo_height(), 300)
        left, right, top, bottom = 92, 30, 28, 62
        plot_width = width - left - right
        plot_height = height - top - bottom
        metric = TREND_METRICS.get(self.trend_metric_var.get(), "revenue")
        if not self.trend_points:
            empty_message = (
                "За выбранный период нет сохраненных отчетов"
                if self.trend_has_history
                else "Импортируйте отчеты, чтобы увидеть динамику"
            )
            canvas.create_text(
                width / 2,
                height / 2,
                text=empty_message,
                fill=self.colors["muted"],
                font=("Segoe UI", 12),
            )
            return
        minimum, maximum = chart_bounds(self.trend_points, metric)
        value_range = maximum - minimum
        for index in range(6):
            ratio = index / 5
            y = top + plot_height * ratio
            value = maximum - value_range * ratio
            canvas.create_line(left, y, width - right, y, fill=self.colors["border"], dash=(2, 4))
            canvas.create_text(
                left - 10,
                y,
                text=_axis_value(value, metric),
                anchor="e",
                fill=self.colors["muted"],
                font=("Segoe UI", 9),
            )
        zero_y = top + (maximum / value_range) * plot_height
        if top <= zero_y <= height - bottom:
            canvas.create_line(left, zero_y, width - right, zero_y, fill=self.colors["muted"], width=1)

        denominator = max(len(self.trend_points) - 1, 1)
        coordinates: list[float] = []
        label_step = max(1, (len(self.trend_points) + 7) // 8)
        for index, point in enumerate(self.trend_points):
            x = left + plot_width * index / denominator if len(self.trend_points) > 1 else left + plot_width / 2
            y = top + (maximum - point.value(metric)) / value_range * plot_height
            coordinates.extend((x, y))
            self.trend_canvas_points.append((x, y, point))
            if index % label_step == 0 or index == len(self.trend_points) - 1:
                canvas.create_text(
                    x,
                    height - bottom + 14,
                    text=_short_period(point.label),
                    anchor="n",
                    fill=self.colors["muted"],
                    font=("Segoe UI", 8),
                    angle=18 if len(self.trend_points) > 5 else 0,
                )
        if len(coordinates) >= 4:
            canvas.create_line(*coordinates, fill=self.colors["accent"], width=3, smooth=False)
        for x, y, point in self.trend_canvas_points:
            color = self.colors["positive"] if point.value(metric) >= 0 else self.colors["negative"]
            canvas.create_oval(x - 5, y - 5, x + 5, y + 5, fill=color, outline=self.colors["surface"], width=2)

    def _trend_hover(self, event) -> None:
        self.trend_canvas.delete("tooltip")
        if not self.trend_canvas_points:
            return
        x, y, point = min(
            self.trend_canvas_points,
            key=lambda item: (item[0] - event.x) ** 2 + (item[1] - event.y) ** 2,
        )
        if (x - event.x) ** 2 + (y - event.y) ** 2 > 225:
            return
        metric = TREND_METRICS.get(self.trend_metric_var.get(), "revenue")
        text = (
            f"Отчет №{self._run_number(point.run_id)}\n"
            f"{point.label}\n{_trend_value(point.value(metric), metric)}"
        )
        text_x = min(max(x + 12, 80), max(self.trend_canvas.winfo_width() - 150, 80))
        text_y = max(y - 58, 8)
        box = self.trend_canvas.create_text(
            text_x,
            text_y,
            text=text,
            anchor="nw",
            fill=self.colors["text"],
            font=("Segoe UI", 9),
            tags="tooltip",
        )
        bounds = self.trend_canvas.bbox(box)
        if bounds:
            background = self.trend_canvas.create_rectangle(
                bounds[0] - 8,
                bounds[1] - 6,
                bounds[2] + 8,
                bounds[3] + 6,
                fill=self.colors["surface_alt"],
                outline=self.colors["border"],
                tags="tooltip",
            )
            self.trend_canvas.tag_lower(background, box)

    def _open_history_run(self, _event=None) -> None:
        selection = self.history_tree.selection()
        if not selection:
            return
        run_id = int(selection[0])
        display = next((key for key, value in self.run_display_to_id.items() if value == run_id), None)
        if display:
            self.run_var.set(display)
        self.overview_selection_explicit = False
        self.select_run(run_id)
        self.notebook.select(self.overview_tab)

    def refresh_comparison(self) -> None:
        first_id = self.run_display_to_id.get(self.compare_first_var.get())
        second_id = self.run_display_to_id.get(self.compare_second_var.get())
        if first_id is None or second_id is None:
            self._clear_comparison("Выберите два сохраненных периода")
            return
        if first_id == second_id:
            self._clear_comparison("Выберите разные периоды")
            return
        comparison = compare_calculations(
            self.db.load_calculation(first_id),
            self.db.load_calculation(second_id),
        )
        self.comparison_kpi_vars["revenue"].set(_comparison_kpi(comparison.revenue, money=True))
        self.comparison_kpi_vars["net_profit"].set(_comparison_kpi(comparison.net_profit, money=True))
        self.comparison_kpi_vars["units"].set(_comparison_kpi(comparison.units, money=False))
        self.comparison_kpi_vars["unallocated"].set(_comparison_kpi(comparison.unallocated, money=True))
        self.comparison_tree.delete(*self.comparison_tree.get_children())
        for row in comparison.products:
            tag = "positive" if row.net_profit.change >= 0 else "negative"
            self.comparison_tree.insert(
                "",
                "end",
                iid=row.article,
                values=(
                    row.article,
                    row.name,
                    _number(row.units.first),
                    _number(row.units.second),
                    _signed_number(row.units.change),
                    _money(row.revenue.first),
                    _money(row.revenue.second),
                    _signed_money(row.revenue.change),
                    _comparison_percent(row.revenue),
                    _money(row.net_profit.first),
                    _money(row.net_profit.second),
                    _signed_money(row.net_profit.change),
                    _comparison_percent(row.net_profit),
                    _percent(row.profitability.first),
                    _percent(row.profitability.second),
                    _signed_percentage_points(row.profitability.change),
                ),
                tags=(tag,),
            )
        self._configure_value_tags(self.comparison_tree)
        self.status_var.set(
            f"Сравнение отчетов №{self._run_number(first_id)} и №{self._run_number(second_id)}"
        )

    def _clear_comparison(self, message: str) -> None:
        if not hasattr(self, "comparison_tree"):
            return
        self.comparison_tree.delete(*self.comparison_tree.get_children())
        for variable in self.comparison_kpi_vars.values():
            variable.set("—")
        self.comparison_tree.insert("", "end", values=("", message))

    def _populate_quality(self, run_id: int | None = None, *, use_current: bool = True) -> None:
        self.quality_tree.delete(*self.quality_tree.get_children())
        target_run_id = self.current_run_id if run_id is None and use_current else run_id
        if target_run_id is None:
            return
        events = self.db.list_quality_events(target_run_id)
        if not events:
            self.quality_tree.insert("", "end", values=("Готово", "Проверки", "Ошибок и предупреждений нет"), tags=("positive",))
        else:
            for event in events:
                tag = "negative" if event["severity"] == "Ошибка" else "warning"
                self.quality_tree.insert("", "end", values=(event["severity"], event["event_type"], event["message"]), tags=(tag,))
        self._configure_value_tags(self.quality_tree)

    def refresh_products(self) -> None:
        if not hasattr(self, "products_tree"):
            return
        self.products_tree.delete(*self.products_tree.get_children())
        products = self.db.list_products()
        query = self.product_search_var.get().strip().casefold() if hasattr(self, "product_search_var") else ""
        status = self.product_status_var.get() if hasattr(self, "product_status_var") else "Все"
        visible = [
            product
            for product in products
            if (
                not query
                or query in product.article.casefold()
                or query in product.name.casefold()
                or query in product.category.casefold()
            )
            and (status == "Все" or (status == "Активные" and product.active) or (status == "Архив" and not product.active))
        ]
        for product in visible:
            self.products_tree.insert(
                "",
                "end",
                iid=product.article,
                values=(
                    product.article, product.name, product.category or CATEGORY_EMPTY,
                    _money(product.total_cost), _money(product.material_cost),
                    _money(product.labor_cost), "Активен" if product.active else "Архив",
                ),
                tags=("" if product.active else "muted",),
            )
        if hasattr(self, "product_count_var"):
            self.product_count_var.set(f"Показано: {len(visible)} из {len(products)}")
        self._configure_value_tags(self.products_tree)

    def _refresh_after_catalog_change(self) -> None:
        self.refresh_products()
        if self.current_run_id is not None:
            self.current_calculation = self.db.load_calculation(self.current_run_id)
            self._refresh_active_report_views()

    def open_cost_catalog_editor(self) -> None:
        dialog = CostCatalogEditorDialog(self, self.db.list_products())
        self.wait_window(dialog)
        if dialog.cancelled:
            return
        changes = build_cost_changes(dialog.products, self.db.product_map(active_only=False))
        changed_products = [change.product for change in changes if change.changed]
        if not changed_products and not dialog.order_changed:
            messagebox.showinfo("Себестоимость", "Изменений нет", parent=self)
            return
        products_to_apply: list[Product] = []
        if changed_products:
            preview = CostImportDialog(self, changes, "Редактор приложения")
            self.wait_window(preview)
            if preview.cancelled:
                return
            products_to_apply = preview.products_to_apply
        try:
            changed = self.db.save_products(products_to_apply, source="Редактор приложения")
            self.db.reorder_products(dialog.article_order)
            self._refresh_after_catalog_change()
            order_text = " Порядок позиций сохранен." if dialog.order_changed else ""
            messagebox.showinfo(
                "Себестоимость сохранена",
                f"Применено изменений: {changed}.\n"
                f"{order_text}\n"
                "Новая себестоимость используется со следующего расчета; числовые показатели истории не меняются. "
                "Категория обновляется во всех сохраненных отчетах.",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Себестоимость", str(exc), parent=self)

    def create_application_backup(self) -> None:
        destination = filedialog.asksaveasfilename(
            title="Создать резервную копию OZ Price Analyzer",
            defaultextension=".ozbackup",
            initialdir=str(self.service.paths["backups"]),
            initialfile=suggested_backup_name(),
            filetypes=[("Резервная копия OZ Price Analyzer", "*.ozbackup")],
            parent=self,
        )
        if not destination:
            return
        self.configure(cursor="watch")
        self.status_var.set("Создание резервной копии…")
        self.update_idletasks()
        try:
            info = create_backup(self.service.paths["root"], destination)
            messagebox.showinfo(
                "Резервная копия создана",
                f"Файл: {info.path}\n\n"
                f"Расчетов: {info.run_count}\n"
                f"Товаров: {info.product_count}\n"
                f"Исходных отчетов: {info.source_count}\n"
                f"Размер исходников: {_file_size(info.source_size)}",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Резервная копия", str(exc), parent=self)
        finally:
            self.configure(cursor="")
            self.status_var.set(self._current_run_status())

    def recalculate_saved_history(self) -> None:
        runs = self.db.list_runs()
        if not runs:
            messagebox.showinfo(
                "Перерасчет истории",
                "В истории пока нет сохраненных отчетов.",
                parent=self,
            )
            return
        confirmed = messagebox.askyesno(
            "Пересчитать историю",
            f"Будут заново рассчитаны все сохраненные отчеты: {len(runs)}.\n\n"
            "Используются сохраненные копии исходных XLSX и PDF. Начисления известных "
            "артикулов будут учтены даже для архивных товаров и товаров без продаж. "
            "Количество продаж будет заново определено по группам «Продажи» и «Возвраты».\n\n"
            "Наименования отчетов, историческая себестоимость и плановые цены сохранятся. "
            "Перед перерасчетом приложение автоматически создаст резервную копию. "
            "Продолжить?",
            parent=self,
        )
        if not confirmed:
            return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = (
            self.service.paths["backups"]
            / f"Автокопия_перед_перерасчетом_{stamp}.ozbackup"
        )
        self.configure(cursor="watch")
        self.status_var.set("Проверка исходных файлов и перерасчет истории…")
        self.update_idletasks()
        try:
            create_backup(self.service.paths["root"], backup_path)
            old_current_id = self.current_run_id
            old_overview_ids = set(self.overview_run_ids)
            result = self.service.recalculate_history()
            self.current_run_id = result.old_to_new.get(old_current_id, old_current_id)
            self.overview_run_ids = {
                result.old_to_new.get(run_id, run_id)
                for run_id in old_overview_ids
            }
            self.current_calculation = None
            self.overview_calculation = None
            self.refresh_all()
            messagebox.showinfo(
                "Перерасчет завершен",
                f"Пересчитано отчетов: {result.replaced_runs}.\n"
                f"Восстановлено товарных строк с движением: {result.recovered_product_rows}.\n"
                f"Изменение количества продаж: {_signed_number(result.units_delta)}.\n"
                f"Изменение чистой прибыли: {_money(result.net_profit_delta)}.\n"
                f"Изменение финансового результата товаров: "
                f"{_money(result.financial_result_delta)}.\n"
                f"Неизвестных пропущенных артикулов: {result.skipped_articles}.\n\n"
                f"Резервная копия до перерасчета:\n{backup_path}",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror(
                "Перерасчет истории",
                f"Историю не удалось пересчитать:\n{exc}\n\n"
                f"Резервная копия до операции: {backup_path}",
                parent=self,
            )
        finally:
            self.configure(cursor="")
            self.status_var.set(self._current_run_status())

    def show_about(self) -> None:
        AboutDialog(self)

    def choose_storage_folder(self) -> None:
        selected = filedialog.askdirectory(
            title="Выберите пустую папку для хранилища OZ Price Analyzer",
            initialdir=str(self.service.paths["root"].parent),
            mustexist=True,
            parent=self,
        )
        if not selected:
            return
        destination = Path(selected).expanduser().resolve()
        source = self.service.paths["root"].resolve()
        if destination == source:
            messagebox.showinfo("Хранилище", "Эта папка уже используется", parent=self)
            return
        confirmed = messagebox.askyesno(
            "Перенести хранилище",
            f"Текущая папка:\n{source}\n\nНовая папка:\n{destination}\n\n"
            "В новую папку будут скопированы история, себестоимость, исходные отчеты, "
            "экспорты и резервные копии. Старая папка останется как дополнительная страховочная копия. "
            "Продолжить?",
            parent=self,
        )
        if not confirmed:
            return
        self.configure(cursor="watch")
        self.status_var.set("Перенос хранилища…")
        self.update_idletasks()
        try:
            result = migrate_storage(source, destination)
            save_storage_location(destination)
            self.service = AppService(destination)
            self.db = self.service.db
            self.current_run_id = None
            self.current_calculation = None
            self.storage_path_var.set(str(destination))
            self._reload_settings_after_restore()
            self.refresh_all()
            messagebox.showinfo(
                "Хранилище перенесено",
                f"Новое хранилище:\n{destination}\n\n"
                f"Расчетов: {result.backup_info.run_count}\n"
                f"Товаров: {result.backup_info.product_count}\n"
                f"Исходных отчетов: {result.backup_info.source_count}\n\n"
                f"Старая папка сохранена:\n{source}",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Перенос хранилища", str(exc), parent=self)
        finally:
            self.configure(cursor="")
            self.status_var.set(self._current_run_status())

    def restore_application_backup(self) -> None:
        source = filedialog.askopenfilename(
            title="Выберите резервную копию OZ Price Analyzer",
            initialdir=str(self.service.paths["backups"]),
            filetypes=[("Резервная копия OZ Price Analyzer", "*.ozbackup")],
            parent=self,
        )
        if not source:
            return
        self.configure(cursor="watch")
        self.status_var.set("Проверка резервной копии…")
        self.update_idletasks()
        try:
            info = inspect_backup(source)
        except Exception as exc:
            self.configure(cursor="")
            self.status_var.set(self._current_run_status())
            messagebox.showerror("Восстановление", str(exc), parent=self)
            return
        self.configure(cursor="")
        confirmed = messagebox.askyesno(
            "Восстановить данные",
            f"Резервная копия: {_backup_timestamp(info.created_at)}\n"
            f"Расчетов: {info.run_count}\n"
            f"Товаров: {info.product_count}\n"
            f"Исходных отчетов: {info.source_count}\n\n"
            "Текущая история будет заменена. Перед заменой приложение автоматически создаст "
            "страховочную копию текущих данных. Продолжить?",
            parent=self,
        )
        if not confirmed:
            self.status_var.set(self._current_run_status())
            return
        self.configure(cursor="watch")
        self.status_var.set("Восстановление истории…")
        self.update_idletasks()
        try:
            result = restore_backup(self.service.paths["root"], source)
            self.service.db = Database(self.service.paths["database"])
            self.db = self.service.db
            self.current_run_id = None
            self.current_calculation = None
            self._reload_settings_after_restore()
            self.refresh_all()
            safety = f"\n\nСтраховочная копия прежних данных:\n{result.safety_backup}" if result.safety_backup else ""
            messagebox.showinfo(
                "Восстановление завершено",
                f"История и настройки восстановлены. Расчетов: {result.info.run_count}.{safety}",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Восстановление", str(exc), parent=self)
        finally:
            self.configure(cursor="")
            self.status_var.set(self._current_run_status())

    def _reload_settings_after_restore(self) -> None:
        self.theme_var.set(THEME_VALUES.get(self.db.get_setting("theme", "system"), "Системная"))
        self.tax_rate_var.set(_plain_number(float(self.db.get_setting("tax_rate", "0.04")) * 100))
        self.warn_realization_var.set(self.db.get_setting("warn_without_realization", "1") == "1")
        self.preview_rows_var.set(self.db.get_setting("preview_rows", "500"))
        if hasattr(self, "storage_path_var"):
            self.storage_path_var.set(str(self.service.paths["root"]))
        self._preview_theme()

    def export_product_catalog(self) -> None:
        destination = filedialog.asksaveasfilename(
            title="Выгрузить справочник себестоимости",
            defaultextension=".xlsx",
            initialdir=str(self.service.paths["exports"]),
            initialfile="Справочник_себестоимости_OZON.xlsx",
            filetypes=[("Книга Excel", "*.xlsx")],
            parent=self,
        )
        if not destination:
            return
        try:
            path = export_cost_catalog(self.db.list_products(), destination)
            messagebox.showinfo(
                "Справочник себестоимости",
                f"Справочник сохранен:\n{path}\n\n"
                "Файл можно использовать как резервную копию или для массового обмена. "
                "Основное редактирование доступно прямо в приложении.",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Экспорт себестоимости", str(exc), parent=self)

    def import_product_catalog(self) -> None:
        source = filedialog.askopenfilename(
            title="Импортировать справочник себестоимости",
            filetypes=[("Книга Excel", "*.xlsx")],
            parent=self,
        )
        if not source:
            return
        try:
            products = read_cost_catalog(source)
            changes = build_cost_changes(products, self.db.product_map(active_only=False))
            dialog = CostImportDialog(self, changes, Path(source).name)
            self.wait_window(dialog)
            if dialog.cancelled:
                return
            changed = self.db.save_products(dialog.products_to_apply, source=f"Импорт: {Path(source).name}")
            self._refresh_after_catalog_change()
            messagebox.showinfo(
                "Импорт себестоимости",
                f"Применено изменений: {changed}.\n"
                "Себестоимость применяется со следующего расчета; числовые показатели истории не меняются. "
                "Категории обновлены во всех сохраненных отчетах.",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Импорт себестоимости", str(exc), parent=self)

    def clear_product_catalog(self) -> None:
        count = len(self.db.list_products())
        if not count:
            messagebox.showinfo("Справочник себестоимости", "Справочник уже пуст", parent=self)
            return
        confirmed = messagebox.askyesno(
            "Очистить справочник?",
            f"Будут удалены все позиции текущего справочника: {count}.\n\n"
            "Сохраненные отчеты и их показатели останутся без изменений. "
            "Новые отчеты нельзя будет полноценно рассчитать, пока вы не загрузите XLSX "
            "или не создадите товары заново.\n\n"
            "Отменить это действие нельзя. Продолжить?",
            icon="warning",
            parent=self,
        )
        if not confirmed:
            return
        removed = self.db.clear_products()
        self._refresh_after_catalog_change()
        messagebox.showinfo(
            "Справочник очищен",
            f"Удалено позиций: {removed}. Сохраненные отчеты не изменены.",
            parent=self,
        )

    def show_cost_history(self) -> None:
        CostHistoryDialog(self, self.db.list_product_cost_history())

    def add_product(self) -> None:
        dialog = ProductDialog(self, title="Новый товар")
        self.wait_window(dialog)
        if dialog.result:
            try:
                self.db.save_product(dialog.result)
                self._refresh_after_catalog_change()
            except Exception as exc:
                messagebox.showerror("Товар", str(exc), parent=self)

    def edit_product(self) -> None:
        selection = self.products_tree.selection()
        if not selection:
            messagebox.showinfo("Товары", "Выберите товар", parent=self)
            return
        product = self.db.product_map(active_only=False)[selection[0]]
        dialog = ProductDialog(self, title="Изменить товар", product=product)
        self.wait_window(dialog)
        if dialog.result:
            try:
                self.db.save_product(dialog.result)
                self._refresh_after_catalog_change()
            except Exception as exc:
                messagebox.showerror("Товар", str(exc), parent=self)

    def toggle_product(self) -> None:
        selection = self.products_tree.selection()
        if not selection:
            return
        product = self.db.product_map(active_only=False)[selection[0]]
        product.active = not product.active
        self.db.save_product(product)
        self._refresh_after_catalog_change()

    def save_settings(self) -> None:
        try:
            tax_percent = _parse_number(self.tax_rate_var.get())
            if tax_percent < 0 or tax_percent > 100:
                raise ValueError
            preview_rows = int(self.preview_rows_var.get())
            if preview_rows < 100 or preview_rows > 5000:
                raise ValueError
        except ValueError:
            messagebox.showerror("Настройки", "Проверьте налоговую ставку и количество строк предпросмотра", parent=self)
            return
        self.db.set_setting("theme", THEME_LABELS[self.theme_var.get()])
        self.db.set_setting("tax_rate", str(tax_percent / 100))
        self.db.set_setting("warn_without_realization", "1" if self.warn_realization_var.get() else "0")
        self.db.set_setting("preview_rows", str(preview_rows))
        self.colors = apply_theme(self, THEME_LABELS[self.theme_var.get()])
        self._configure_help_text_theme()
        messagebox.showinfo("Настройки", "Настройки сохранены", parent=self)

    def _preview_theme(self, _event=None) -> None:
        self.colors = apply_theme(self, THEME_LABELS[self.theme_var.get()])
        self._configure_help_text_theme()
        for tree in (
            self.overview_tree,
            self.breakdown_tree,
            self.scenario_tree,
            self.history_tree,
            self.quality_tree,
            self.trend_tree,
            self.comparison_tree,
            self.products_tree,
        ):
            self._configure_value_tags(tree)
        self._draw_trend_chart()

    def import_reports(self) -> None:
        if self.import_in_progress:
            messagebox.showinfo("Импорт отчетов", "Проверка файлов уже выполняется", parent=self)
            return
        paths = filedialog.askopenfilenames(
            title="Выберите отчеты Ozon",
            filetypes=[
                ("Отчеты Ozon", "*.xlsx *.pdf"),
                ("Отчеты Excel", "*.xlsx"),
                ("Акты Ozon PDF", "*.pdf"),
            ],
            parent=self,
        )
        if not paths:
            return
        self.configure(cursor="watch")
        self.status_var.set("Проверка исходных файлов…")
        self.update_idletasks()
        self.import_in_progress = True
        threading.Thread(target=self._prepare_import_worker, args=(list(paths),), daemon=True).start()
        self.after(100, self._poll_import_queue)

    def _prepare_import_worker(self, paths: list[str]) -> None:
        try:
            batch = self.service.prepare_import(list(paths))
            self.import_queue.put((batch, None))
        except Exception as exc:
            self.import_queue.put((None, exc))

    def _poll_import_queue(self) -> None:
        try:
            batch, error = self.import_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_import_queue)
            return
        self._complete_import_ui(batch, error)

    def _complete_import_ui(self, batch: ImportBatch | None, error: Exception | None) -> None:
        try:
            if error is not None:
                raise error
            if batch is None or not batch.sessions:
                raise RuntimeError("Не удалось подготовить импорт")
            sessions = batch.sessions
            if any(not session.sources or not session.has_accrual for session in sessions):
                raise ValueError("Не найден отчет по начислениям")

            warnings_by_session = [session.realization_period_warnings() for session in sessions]
            all_period_warnings = [
                f"{_session_period(session)}: {warning}"
                for session, warnings in zip(sessions, warnings_by_session)
                for warning in warnings
            ]
            if all_period_warnings:
                dialog = RealizationPeriodDialog(self, all_period_warnings)
                self.wait_window(dialog)
                if not dialog.confirmed:
                    return

            replacements_by_session = [
                self.service.replacement_run_ids(session) for session in sessions
            ]
            for session, replace_run_ids in zip(sessions, replacements_by_session):
                if not replace_run_ids:
                    continue
                runs_by_id = {run.id: run for run in self.db.list_runs()}
                dialog = ReplacePeriodDialog(
                    self,
                    period=_session_period(session),
                    runs=[runs_by_id[run_id] for run_id in replace_run_ids if run_id in runs_by_id],
                    run_numbers=self.run_number_by_id,
                )
                self.wait_window(dialog)
                if not dialog.confirmed:
                    return

            sessions_without_realization = [
                session for session in sessions if not session.has_realization
            ]
            if (
                sessions_without_realization
                and self.db.get_setting("warn_without_realization", "1") == "1"
            ):
                periods = "\n".join(
                    f"• {_session_period(session)}" for session in sessions_without_realization
                )
                if not messagebox.askyesno(
                    "Нет отчета о выкупленных товарах",
                    "Для следующих периодов не выбран RealizationReportCIS:\n"
                    f"{periods}\n\n"
                    "Продолжить без отчета по выкупам? Если такие продажи были, "
                    "выручка этих отчетов будет неполной.",
                    parent=self,
                ):
                    return

            batch.unknown_products = discover_unknown_products(
                batch.sources,
                self.db.product_map(active_only=False),
            )
            created: list[Product] = []
            skipped: set[str] = set()
            if batch.unknown_products:
                dialog = UnknownProductsDialog(self, batch.unknown_products)
                self.wait_window(dialog)
                if dialog.cancelled:
                    return
                created = dialog.created_products
                skipped = dialog.skipped_articles

            calculations = self.service.complete_import_batch(
                sessions,
                created_products=created,
                skipped_articles=skipped,
                replace_run_ids_by_session=replacements_by_session,
                source_period_warnings_by_session=warnings_by_session,
            )

            self.current_run_id = calculations[-1].run_id
            self.refresh_all()
            updated_count = sum(bool(run_ids) for run_ids in replacements_by_session)
            created_count = len(calculations) - updated_count
            details = "\n".join(
                f"• {_calculation_period(calculation)} — "
                f"{_realization_file_count(calculation)}; "
                f"выручка {_money(calculation.realization_revenue)}"
                for calculation in calculations
            )
            result_message = (
                f"Обработано отчетов по начислениям: {len(calculations)}.\n"
                f"Создано новых отчетов: {created_count}. Обновлено: {updated_count}.\n\n"
                f"{details}"
            )
            messagebox.showinfo(
                "Пакет отчетов обработан" if len(calculations) > 1 else "Расчет готов",
                result_message,
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Импорт отчетов", str(exc), parent=self)
        finally:
            self.import_in_progress = False
            self.configure(cursor="")
            self.status_var.set(self._current_run_status())

    def export_current_run(self) -> None:
        if self.current_run_id is None or self.current_calculation is None:
            messagebox.showinfo("Экспорт", "Сначала импортируйте или выберите расчет", parent=self)
            return
        export_overview = (
            self.notebook.select() == str(self.overview_tab)
            and self.overview_selection_explicit
            and self.overview_calculation is not None
        )
        calculation = self.overview_calculation if export_overview else self.current_calculation
        destination = filedialog.asksaveasfilename(
            title="Сохранить итоговый отчет",
            defaultextension=".xlsx",
            initialdir=str(self.service.paths["exports"]),
            initialfile=suggested_export_name(calculation),
            filetypes=[("Книга Excel", "*.xlsx")],
            parent=self,
        )
        if not destination:
            return
        try:
            path = (
                export_calculation(calculation, destination)
                if export_overview
                else export_run(self.db, self.current_run_id, destination)
            )
            description = "Сводный отчет за выбранный период" if export_overview else "Отчет"
            messagebox.showinfo("Экспорт", f"{description} сохранен:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Экспорт", str(exc), parent=self)

    def _configure_value_tags(self, tree: ttk.Treeview) -> None:
        palette = self.colors
        tree.tag_configure("negative", foreground=palette["negative"])
        tree.tag_configure("positive", foreground=palette["positive"])
        tree.tag_configure("warning", foreground=palette["warning"])
        tree.tag_configure("total", background=palette["surface_alt"], foreground=palette["text"])
        tree.tag_configure("muted", foreground=palette["muted"])


class OverviewReportSelectionDialog(tk.Toplevel):
    def __init__(
        self,
        parent: OZPriceAnalyzerApp,
        runs: list[RunSummary],
        selected_run_ids: set[int],
    ):
        super().__init__(parent)
        self.title("Отчеты для итогового обзора")
        self.geometry("1040x650")
        self.minsize(850, 520)
        self.transient(parent)
        self.grab_set()
        self.confirmed = False
        self.runs = runs
        self.selected_run_ids = {run.id for run in runs if run.id in selected_run_ids}
        self.run_by_id = {run.id: run for run in runs}
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        ttk.Label(
            self,
            text="Сформировать обзор за несколько месяцев",
            style="Section.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=22, pady=(20, 3))
        ttk.Label(
            self,
            text=(
                "Отметьте нужные отчеты или выберите целый год. Показатели товаров, "
                "себестоимость, налог и нераспределенные суммы будут объединены за весь период."
            ),
            style="Muted.TLabel",
            wraplength=950,
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 12))

        controls = ttk.Frame(self)
        controls.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 10))
        controls.columnconfigure(7, weight=1)
        ttk.Label(controls, text="Год:").grid(row=0, column=0, padx=(0, 6))
        years = sorted({_run_year(run) for run in runs if _run_year(run) is not None})
        self.year_var = tk.StringVar(value=str(years[-1]) if years else "Все годы")
        self.year_combo = ttk.Combobox(
            controls,
            textvariable=self.year_var,
            values=["Все годы"] + [str(year) for year in years],
            state="readonly",
            width=14,
        )
        self.year_combo.grid(row=0, column=1, padx=(0, 8))
        ttk.Button(controls, text="Выбрать год", command=self._select_year).grid(
            row=0, column=2, padx=(0, 12)
        )
        ttk.Button(controls, text="Выбрать все", command=self._select_all).grid(
            row=0, column=3, padx=4
        )
        ttk.Button(controls, text="Снять все", command=self._clear).grid(
            row=0, column=4, padx=4
        )
        ttk.Label(
            controls,
            text="Щелчок по строке включает или исключает отчет",
            style="Muted.TLabel",
        ).grid(row=0, column=7, sticky="e")

        container = ttk.Frame(self)
        container.grid(row=3, column=0, sticky="nsew", padx=22)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        columns = ("selected", "number", "period", "name", "files", "revenue")
        self.tree = ttk.Treeview(container, columns=columns, show="headings", height=14)
        headings = ("Выбран", "№", "Период", "Наименование", "Файлов", "Выручка")
        widths = (85, 55, 210, 430, 80, 150)
        for column, heading, width in zip(columns, headings, widths):
            self.tree.heading(column, text=heading)
            self.tree.column(
                column,
                width=width,
                minwidth=50,
                stretch=column == "name",
                anchor="w" if column in {"period", "name"} else "center",
            )
        yscroll = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<ButtonRelease-1>", self._toggle_clicked)
        self.tree.bind("<space>", self._toggle_focused)
        self.tree.tag_configure("selected", foreground=parent.colors["positive"])

        footer = ttk.Frame(self, padding=(22, 12, 22, 18))
        footer.grid(row=4, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        self.summary_var = tk.StringVar()
        ttk.Label(footer, textvariable=self.summary_var, style="Section.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(footer, text="Отмена", command=self.destroy).grid(
            row=0, column=1, padx=4
        )
        ttk.Button(
            footer,
            text="Показать итоговый обзор",
            style="Accent.TButton",
            command=self._confirm,
        ).grid(row=0, column=2, padx=4)
        self.bind("<Escape>", lambda _event: self.destroy())
        self._refresh_rows()

    def _refresh_rows(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for number, run in enumerate(self.runs, start=1):
            selected = run.id in self.selected_run_ids
            self.tree.insert(
                "",
                "end",
                iid=str(run.id),
                values=(
                    "✓" if selected else "",
                    number,
                    _period_text(run.period_start, run.period_end),
                    run.report_name,
                    run.source_count,
                    _money(run.revenue),
                ),
                tags=("selected",) if selected else (),
            )
        chosen = [run for run in self.runs if run.id in self.selected_run_ids]
        if chosen:
            starts = [value for run in chosen if (value := _summary_date(run.period_start))]
            ends = [value for run in chosen if (value := _summary_date(run.period_end))]
            period = (
                f"{min(starts):%d.%m.%Y}–{max(ends):%d.%m.%Y}"
                if starts and ends
                else "не определен"
            )
            self.summary_var.set(
                f"Выбрано: {_russian_report_count(len(chosen))} · итоговый период {period}"
            )
        else:
            self.summary_var.set("Отчеты не выбраны")

    def _toggle_clicked(self, event) -> None:
        item_id = self.tree.identify_row(event.y)
        if item_id:
            self._toggle(int(item_id))

    def _toggle_focused(self, _event=None) -> str:
        item_id = self.tree.focus()
        if item_id:
            self._toggle(int(item_id))
        return "break"

    def _toggle(self, run_id: int) -> None:
        if run_id in self.selected_run_ids:
            self.selected_run_ids.remove(run_id)
        else:
            self.selected_run_ids.add(run_id)
        self._refresh_rows()
        if self.tree.exists(str(run_id)):
            self.tree.focus(str(run_id))
            self.tree.see(str(run_id))

    def _select_year(self) -> None:
        value = self.year_var.get()
        if value == "Все годы":
            self.selected_run_ids = {run.id for run in self.runs}
        else:
            year = int(value)
            self.selected_run_ids = {
                run.id for run in self.runs if _run_year(run) == year
            }
        self._refresh_rows()

    def _select_all(self) -> None:
        self.selected_run_ids = {run.id for run in self.runs}
        self._refresh_rows()

    def _clear(self) -> None:
        self.selected_run_ids.clear()
        self._refresh_rows()

    def _confirm(self) -> None:
        if not self.selected_run_ids:
            messagebox.showwarning(
                "Период обзора",
                "Выберите хотя бы один отчет.",
                parent=self,
            )
            return
        self.confirmed = True
        self.destroy()


class RealizationPeriodDialog(tk.Toplevel):
    def __init__(self, parent: OZPriceAnalyzerApp, warnings: list[str]):
        super().__init__(parent)
        self.title("Периоды отчетов не совпадают")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.configure(background=parent.colors["window"])
        self.confirmed = False

        ttk.Label(
            self,
            text="Проверьте период отчета о выкупленных товарах",
            style="Section.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(22, 4))
        ttk.Label(
            self,
            text=(
                "Отчет по выкупам должен относиться к тому же календарному месяцу "
                "и году, что и отчет по начислениям. Найдены расхождения:"
            ),
            justify="left",
            wraplength=730,
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 10))
        ttk.Label(
            self,
            text="\n".join(f"• {message}" for message in warnings),
            justify="left",
            wraplength=730,
            style="Warning.TLabel",
        ).grid(row=2, column=0, sticky="w", padx=24, pady=(0, 12))
        ttk.Label(
            self,
            text=(
                "Если продолжить, выручка из указанных RealizationReportCIS войдет в расчет, "
                "а расхождение будет записано в «Контроль качества»."
            ),
            justify="left",
            wraplength=730,
            style="Muted.TLabel",
        ).grid(row=3, column=0, sticky="w", padx=24, pady=(0, 18))

        buttons = ttk.Frame(self)
        buttons.grid(row=4, column=0, sticky="e", padx=20, pady=(0, 20))
        ttk.Button(buttons, text="Отменить импорт", command=self.destroy).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(
            buttons,
            text="Продолжить с несовпадающим периодом",
            style="Accent.TButton",
            command=self._confirm,
        ).grid(row=0, column=1, padx=4)
        self.bind("<Escape>", lambda _event: self.destroy())

    def _confirm(self) -> None:
        self.confirmed = True
        self.destroy()


class ReplacePeriodDialog(tk.Toplevel):
    def __init__(
        self,
        parent: OZPriceAnalyzerApp,
        period: str,
        runs: list[RunSummary],
        run_numbers: dict[int, int],
    ):
        super().__init__(parent)
        self.title("Отчет за этот период уже есть")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.configure(background=parent.colors["window"])
        self.confirmed = False

        ttk.Label(self, text="Обновить отчет за период?", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=24, pady=(22, 4)
        )
        ttk.Label(
            self,
            text=f"Период: {period}",
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 8))
        names = "\n".join(
            f"• №{run_numbers.get(run.id, '—')} · {run.report_name}"
            for run in runs
        )
        ttk.Label(
            self,
            text=(
                "В истории уже есть:\n"
                f"{names}\n\n"
                "При обновлении новый отчет будет сначала рассчитан и сохранен. "
                "Только после этого прежний отчет будет удален. "
                "Если возникнет ошибка, старый отчет останется без изменений."
            ),
            justify="left",
            wraplength=650,
        ).grid(row=2, column=0, sticky="w", padx=24, pady=(0, 18))

        buttons = ttk.Frame(self)
        buttons.grid(row=3, column=0, sticky="e", padx=20, pady=(0, 20))
        ttk.Button(buttons, text="Прервать импорт", command=self.destroy).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(
            buttons,
            text="Обновить отчет за период",
            style="Accent.TButton",
            command=self._confirm,
        ).grid(row=0, column=1, padx=4)
        self.bind("<Escape>", lambda _event: self.destroy())

    def _confirm(self) -> None:
        self.confirmed = True
        self.destroy()


class HistoryYearFilterDialog(tk.Toplevel):
    def __init__(
        self,
        parent: OZPriceAnalyzerApp,
        years: list[int],
        selected_years: set[int] | None,
        *,
        title: str = "Фильтр истории по годам",
    ):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.configure(background=parent.colors["window"])
        self.cancelled = True
        self.years = years
        self.selected_years: set[int] | None = selected_years
        initially_selected = set(years) if selected_years is None else set(selected_years)
        self.year_vars = {
            year: tk.BooleanVar(value=year in initially_selected)
            for year in years
        }

        ttk.Label(self, text="Какие годы показывать", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=22, pady=(20, 2)
        )
        ttk.Label(
            self,
            text="Выберите один или несколько лет. «Все годы» включает всю историю.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 12))

        years_frame = ttk.LabelFrame(self, text="Годы", padding=(14, 10))
        years_frame.grid(row=2, column=0, sticky="ew", padx=22)
        for index, year in enumerate(years):
            ttk.Checkbutton(
                years_frame,
                text=str(year),
                variable=self.year_vars[year],
            ).grid(row=index // 4, column=index % 4, sticky="w", padx=(0, 24), pady=4)

        selection_buttons = ttk.Frame(self)
        selection_buttons.grid(row=3, column=0, sticky="w", padx=22, pady=(10, 0))
        ttk.Button(selection_buttons, text="Все годы", command=lambda: self._set_all(True)).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(selection_buttons, text="Снять выбор", command=lambda: self._set_all(False)).grid(
            row=0, column=1
        )

        buttons = ttk.Frame(self, padding=(20, 14))
        buttons.grid(row=4, column=0, sticky="e")
        ttk.Button(buttons, text="Отмена", command=self.destroy).grid(row=0, column=0, padx=4)
        ttk.Button(buttons, text="Применить", style="Accent.TButton", command=self._finish).grid(
            row=0, column=1, padx=4
        )
        self.bind("<Escape>", lambda _event: self.destroy())

    def _set_all(self, selected: bool) -> None:
        for variable in self.year_vars.values():
            variable.set(selected)

    def _finish(self) -> None:
        selected = {year for year, variable in self.year_vars.items() if variable.get()}
        if not selected:
            messagebox.showerror("Фильтр по годам", "Выберите хотя бы один год", parent=self)
            return
        self.selected_years = None if selected == set(self.years) else selected
        self.cancelled = False
        self.destroy()


class AboutDialog(tk.Toplevel):
    def __init__(self, parent: OZPriceAnalyzerApp):
        super().__init__(parent)
        self.title("О программе")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.configure(background=parent.colors["window"])
        self.columnconfigure(0, weight=1)

        ttk.Label(self, text=APP_TITLE, style="Title.TLabel").grid(
            row=0, column=0, sticky="w", padx=24, pady=(22, 2)
        )
        ttk.Label(self, text=f"Версия {APP_VERSION}", style="Section.TLabel").grid(
            row=1, column=0, sticky="w", padx=24
        )
        ttk.Label(
            self,
            text="Локальный анализ отчетов Ozon, контроль начислений, история и сценарии доходности.",
            style="Muted.TLabel",
            wraplength=560,
            justify="left",
        ).grid(row=2, column=0, sticky="w", padx=24, pady=(10, 14))

        details = ttk.LabelFrame(self, text="Сведения", padding=(14, 10))
        details.grid(row=3, column=0, sticky="ew", padx=24)
        details.columnconfigure(1, weight=1)
        build_type = "Автономная Windows-сборка" if getattr(sys, "frozen", False) else "Запуск из Python"
        for row, (label, value) in enumerate(
            [
                ("Тип запуска", build_type),
                ("Хранилище данных", str(parent.service.paths["root"])),
                ("Репозиторий", "github.com/otdelvsego-spec/OZPriceAnalyzer"),
            ]
        ):
            ttk.Label(details, text=f"{label}:").grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=3)
            ttk.Label(details, text=value, style="Muted.TLabel", wraplength=430).grid(
                row=row, column=1, sticky="w", pady=3
            )

        ttk.Label(
            self,
            text="Microsoft Excel не требуется. Все рабочие данные остаются на этом компьютере.",
            style="Muted.TLabel",
        ).grid(row=4, column=0, sticky="w", padx=24, pady=(12, 4))

        buttons = ttk.Frame(self, padding=(20, 14))
        buttons.grid(row=5, column=0, sticky="e")
        ttk.Button(
            buttons,
            text="Открыть хранилище",
            command=lambda: _open_path(parent.service.paths["root"]),
        ).grid(row=0, column=0, padx=4)
        ttk.Button(
            buttons,
            text="Открыть GitHub",
            command=lambda: webbrowser.open("https://github.com/otdelvsego-spec/OZPriceAnalyzer"),
        ).grid(row=0, column=1, padx=4)
        ttk.Button(buttons, text="Закрыть", style="Accent.TButton", command=self.destroy).grid(
            row=0, column=2, padx=4
        )
        self.bind("<Escape>", lambda _event: self.destroy())


class CostCatalogEditorDialog(tk.Toplevel):
    def __init__(self, parent: OZPriceAnalyzerApp, products: list[Product]):
        super().__init__(parent)
        self.title("Редактор товаров и себестоимости")
        self.geometry("1280x760")
        self.minsize(1020, 650)
        self.transient(parent)
        self.grab_set()
        self.cancelled = True
        self.products: list[Product] = []
        self.article_order: list[str] = []
        self.order_changed = False
        self.product_map = {
            product.article: Product(
                article=product.article,
                name=product.name,
                material_cost=product.material_cost,
                labor_cost=product.labor_cost,
                active=product.active,
                sort_order=product.sort_order,
                category=product.category,
            )
            for product in products
        }
        self.product_order = [product.article for product in products]
        self.original_order = list(self.product_order)
        self.original_articles = set(self.product_map)
        self.current_article: str | None = None
        self.loading = False
        self.dirty = False

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        ttk.Label(self, text="Товары и себестоимость", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=20, pady=(18, 2)
        )
        ttk.Label(
            self,
            text=(
                "Заполняйте справочник прямо здесь. Полная себестоимость состоит из материала и трудозатрат. "
                "Перед окончательным сохранением приложение покажет все изменения."
            ),
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 10))

        controls = ttk.Frame(self)
        controls.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 8))
        controls.columnconfigure(2, weight=1)
        ttk.Label(controls, text="Поиск:").grid(row=0, column=0, padx=(0, 6))
        self.search_var = tk.StringVar()
        ttk.Entry(controls, textvariable=self.search_var, width=32).grid(row=0, column=1, padx=(0, 12))
        self.search_var.trace_add("write", lambda *_args: self._refresh_tree())
        self.count_var = tk.StringVar()
        ttk.Label(controls, textvariable=self.count_var, style="Muted.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Button(controls, text="↑ Вверх", command=lambda: self._move_selected(-1)).grid(
            row=0, column=3, padx=3
        )
        ttk.Button(controls, text="↓ Вниз", command=lambda: self._move_selected(1)).grid(
            row=0, column=4, padx=3
        )
        ttk.Button(controls, text="Новая позиция", style="Accent.TButton", command=self._new_product).grid(
            row=0, column=5, padx=(8, 0)
        )

        container = ttk.Frame(self)
        container.grid(row=3, column=0, sticky="nsew", padx=20)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        columns = ("article", "name", "category", "total", "material", "labor", "status")
        self.tree = ttk.Treeview(container, columns=columns, show="headings", selectmode="browse")
        headings = ["Артикул", "Наименование", "Категория", "Полная себестоимость", "Материал", "Трудозатраты", "Статус"]
        widths = [150, 300, 220, 170, 150, 150, 100]
        for column, heading, width in zip(columns, headings, widths):
            self.tree.heading(column, text=heading)
            self.tree.column(
                column,
                width=width,
                minwidth=80,
                stretch=False,
                anchor="w" if column in {"article", "name", "category", "status"} else "e",
            )
        xscroll = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        yscroll = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda _event: self.name_entry.focus_set())
        self.tree.tag_configure("muted", foreground=parent.colors["muted"])

        editor = ttk.LabelFrame(self, text="Редактирование выбранной позиции", padding=(14, 10))
        editor.grid(row=4, column=0, sticky="ew", padx=20, pady=(12, 0))
        editor.columnconfigure(3, weight=1)
        self.article_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.category_var = tk.StringVar()
        self.total_var = tk.StringVar()
        self.labor_var = tk.StringVar(value="0")
        self.material_var = tk.StringVar(value="—")
        self.active_var = tk.BooleanVar(value=True)
        ttk.Label(editor, text="Артикул:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self.article_entry = ttk.Entry(editor, textvariable=self.article_var, width=22)
        self.article_entry.grid(row=0, column=1, sticky="w", padx=(0, 18), pady=4)
        ttk.Label(editor, text="Наименование:").grid(row=0, column=2, sticky="w", padx=(0, 6), pady=4)
        self.name_entry = ttk.Entry(editor, textvariable=self.name_var)
        self.name_entry.grid(row=0, column=3, sticky="ew", padx=(0, 18), pady=4)
        ttk.Checkbutton(editor, text="Активен", variable=self.active_var).grid(row=0, column=4, sticky="w", pady=4)

        ttk.Label(editor, text="Категория:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(editor, textvariable=self.category_var).grid(
            row=1, column=1, columnspan=3, sticky="ew", padx=(0, 18), pady=4
        )
        ttk.Label(editor, text="Полная себестоимость, руб.:").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(editor, textvariable=self.total_var, width=22).grid(row=2, column=1, sticky="w", padx=(0, 18), pady=4)
        ttk.Label(editor, text="Трудозатраты, руб.:").grid(row=2, column=2, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(editor, textvariable=self.labor_var, width=18).grid(row=2, column=3, sticky="w", pady=4)
        ttk.Label(editor, text="Материал рассчитывается автоматически:").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        ttk.Label(editor, textvariable=self.material_var, style="Section.TLabel").grid(
            row=3, column=2, sticky="w", pady=(4, 0)
        )
        ttk.Button(editor, text="Применить в таблицу", command=self._commit_current).grid(
            row=3, column=4, sticky="e", pady=(4, 0)
        )

        for variable in (self.article_var, self.name_var, self.category_var, self.total_var, self.labor_var):
            variable.trace_add("write", self._field_changed)
        self.active_var.trace_add("write", self._field_changed)

        buttons = ttk.Frame(self, padding=(20, 14))
        buttons.grid(row=5, column=0, sticky="e")
        ttk.Button(buttons, text="Отмена", command=self.destroy).grid(row=0, column=0, padx=4)
        ttk.Button(
            buttons,
            text="Проверить и сохранить справочник",
            style="Accent.TButton",
            command=self._finish,
        ).grid(row=0, column=1, padx=4)
        self.bind("<Control-s>", lambda _event: self._finish())
        self.bind("<Escape>", lambda _event: self.destroy())

        self._refresh_tree()
        first = next(iter(self.product_map), None)
        if first:
            self._select_article(first)
        else:
            self._new_product()

    def _field_changed(self, *_args) -> None:
        if self.loading:
            return
        self.dirty = True
        try:
            total = _parse_number(self.total_var.get())
            labor = _parse_number(self.labor_var.get() or "0")
            self.material_var.set(_money(total - labor) if total >= labor >= 0 else "Проверьте значения")
        except ValueError:
            self.material_var.set("—")

    def _refresh_tree(self) -> None:
        if not hasattr(self, "tree"):
            return
        query = self.search_var.get().strip().casefold()
        selected = self.current_article
        self.loading = True
        try:
            self.tree.delete(*self.tree.get_children())
            visible = [
                self.product_map[article]
                for article in self.product_order
                if article in self.product_map
                and (
                    not query
                    or query in article.casefold()
                    or query in self.product_map[article].name.casefold()
                    or query in self.product_map[article].category.casefold()
                )
            ]
            for product in visible:
                self.tree.insert(
                    "",
                    "end",
                    iid=product.article,
                    values=(
                        product.article,
                        product.name,
                        product.category or CATEGORY_EMPTY,
                        _money(product.total_cost),
                        _money(product.material_cost),
                        _money(product.labor_cost),
                        "Активен" if product.active else "Архив",
                    ),
                    tags=("" if product.active else "muted",),
                )
            self.count_var.set(f"Показано: {len(visible)} из {len(self.product_map)}")
            if selected and self.tree.exists(selected):
                self.tree.selection_set(selected)
                self.tree.focus(selected)
        finally:
            self.loading = False

    def _move_selected(self, direction: int) -> None:
        if self.search_var.get().strip():
            messagebox.showinfo(
                "Порядок товаров",
                "Очистите поиск, чтобы менять позиции в полном списке.",
                parent=self,
            )
            return
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Порядок товаров", "Выберите позицию в таблице", parent=self)
            return
        if self.dirty and not self._commit_current():
            return
        article = selection[0]
        index = self.product_order.index(article)
        target = index + direction
        if target < 0 or target >= len(self.product_order):
            return
        self.product_order[index], self.product_order[target] = (
            self.product_order[target],
            self.product_order[index],
        )
        self._refresh_tree()
        self._select_article(article)

    def _on_select(self, _event=None) -> None:
        if self.loading:
            return
        selection = self.tree.selection()
        if not selection:
            return
        target = selection[0]
        if target == self.current_article:
            return
        if self.dirty:
            answer = messagebox.askyesnocancel(
                "Несохраненная строка",
                "Применить изменения текущей строки перед переходом к другой позиции?",
                parent=self,
            )
            if answer is None:
                self._select_article(self.current_article)
                return
            if answer and not self._commit_current():
                self._select_article(self.current_article)
                return
        self._load_product(target)

    def _select_article(self, article: str | None) -> None:
        if not article:
            self.loading = True
            try:
                self.tree.selection_remove(*self.tree.selection())
            finally:
                self.loading = False
            return
        if not self.tree.exists(article):
            return
        self.loading = True
        try:
            self.tree.selection_set(article)
            self.tree.focus(article)
            self.tree.see(article)
        finally:
            self.loading = False
        self._load_product(article)

    def _load_product(self, article: str) -> None:
        product = self.product_map[article]
        self.loading = True
        try:
            self.current_article = article
            self.article_var.set(product.article)
            self.name_var.set(product.name)
            self.category_var.set(product.category)
            self.total_var.set(_plain_number(product.total_cost))
            self.labor_var.set(_plain_number(product.labor_cost))
            self.material_var.set(_money(product.material_cost))
            self.active_var.set(product.active)
            self.article_entry.configure(state="disabled" if article in self.original_articles else "normal")
            self.dirty = False
        finally:
            self.loading = False

    def _new_product(self) -> None:
        if self.dirty:
            if not self._commit_current():
                return
        self.loading = True
        try:
            self.tree.selection_remove(*self.tree.selection())
            self.current_article = None
            self.article_var.set("")
            self.name_var.set("")
            self.category_var.set("")
            self.total_var.set("")
            self.labor_var.set("0")
            self.material_var.set("—")
            self.active_var.set(True)
            self.article_entry.configure(state="normal")
            self.dirty = False
        finally:
            self.loading = False
        self.article_entry.focus_set()

    def _commit_current(self) -> bool:
        entry = CostEditorEntry(
            article=self.article_var.get(),
            name=self.name_var.get(),
            total_cost=self.total_var.get(),
            labor_cost=self.labor_var.get(),
            active=self.active_var.get(),
            category=self.category_var.get(),
            row_number=(self.product_order.index(self.current_article) + 1)
            if self.current_article in self.product_order
            else len(self.product_map) + 1,
        )
        try:
            product = build_products_from_editor_entries([entry])[0]
            if product.article != self.current_article and product.article in self.product_map:
                raise ValueError(f"Артикул {product.article} уже есть в справочнике")
        except Exception as exc:
            messagebox.showerror("Себестоимость", str(exc), parent=self)
            return False
        previous_article = self.current_article
        if previous_article and previous_article != product.article and previous_article not in self.original_articles:
            self.product_map.pop(previous_article, None)
            if previous_article in self.product_order:
                self.product_order[self.product_order.index(previous_article)] = product.article
        elif product.article not in self.product_order:
            self.product_order = insert_at_group_end(self.product_order, product.article)
        self.product_map[product.article] = product
        self.current_article = product.article
        self.dirty = False
        self._refresh_tree()
        self._select_article(product.article)
        return True

    def _finish(self) -> None:
        if self.dirty and not self._commit_current():
            return
        try:
            entries = [
                CostEditorEntry(
                    article=product.article,
                    name=product.name,
                    total_cost=product.total_cost,
                    labor_cost=product.labor_cost,
                    active=product.active,
                    row_number=index,
                    category=product.category,
                )
                for index, product in enumerate(
                    (self.product_map[article] for article in self.product_order), start=1
                )
            ]
            self.products = build_products_from_editor_entries(entries)
            self.article_order = [product.article for product in self.products]
            for index, product in enumerate(self.products, start=1):
                product.sort_order = index
            self.order_changed = self.article_order != self.original_order
        except Exception as exc:
            messagebox.showerror("Себестоимость", str(exc), parent=self)
            return
        self.cancelled = False
        self.destroy()


class CostImportDialog(tk.Toplevel):
    def __init__(self, parent: OZPriceAnalyzerApp, changes: list[CostChange], source_name: str):
        super().__init__(parent)
        self.title("Предварительная проверка себестоимости")
        self.geometry("1260x650")
        self.minsize(980, 520)
        self.transient(parent)
        self.grab_set()
        self.cancelled = True
        self.changes = changes
        self.products_to_apply = [change.product for change in changes if change.changed]
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        ttk.Label(self, text="Проверьте изменения перед применением", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=20, pady=(18, 2)
        )
        changed_count = len(self.products_to_apply)
        new_count = sum(change.status == "Новая позиция" for change in changes)
        ttk.Label(
            self,
            text=(
                f"Источник: {source_name} · строк: {len(changes)} · изменений: {changed_count} · новых товаров: {new_count}. "
                "Старые отчеты и их себестоимость останутся без изменений."
            ),
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 10))
        ttk.Label(
            self,
            text="Зеленым отмечены новые и измененные позиции, серым — строки без изменений.",
            style="Muted.TLabel",
        ).grid(row=2, column=0, sticky="w", padx=20, pady=(0, 8))

        container = ttk.Frame(self)
        container.grid(row=3, column=0, sticky="nsew", padx=20)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        columns = (
            "article", "name", "old_category", "new_category", "old_total", "new_total", "change",
            "old_labor", "new_labor", "active", "status",
        )
        self.tree = ttk.Treeview(container, columns=columns, show="headings")
        headings = [
            "Артикул", "Наименование", "Старая категория", "Новая категория", "Старая с/с", "Новая с/с", "Изменение",
            "Старые трудозатраты", "Новые трудозатраты", "Активен", "Действие",
        ]
        widths = [140, 260, 180, 180, 130, 130, 130, 160, 160, 90, 150]
        for column, heading, width in zip(columns, headings, widths):
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, minwidth=80, stretch=False, anchor="w" if column in {"article", "name", "old_category", "new_category", "status"} else "e")
        xscroll = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        yscroll = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        for change in changes:
            previous = change.previous
            tag = "changed" if change.changed else "muted"
            self.tree.insert(
                "",
                "end",
                values=(
                    change.product.article,
                    change.product.name,
                    (previous.category or CATEGORY_EMPTY) if previous else "—",
                    change.product.category or CATEGORY_EMPTY,
                    _money(previous.total_cost) if previous else "—",
                    _money(change.product.total_cost),
                    _signed_money(change.total_change) if change.total_change is not None else "Новая",
                    _money(previous.labor_cost) if previous else "—",
                    _money(change.product.labor_cost),
                    "Да" if change.product.active else "Нет",
                    change.status,
                ),
                tags=(tag,),
            )
        self.tree.tag_configure("changed", foreground=parent.colors["positive"])
        self.tree.tag_configure("muted", foreground=parent.colors["muted"])

        buttons = ttk.Frame(self, padding=(20, 14))
        buttons.grid(row=4, column=0, sticky="e")
        ttk.Button(buttons, text="Отмена", command=self.destroy).grid(row=0, column=0, padx=4)
        apply_button = ttk.Button(buttons, text="Применить изменения", style="Accent.TButton", command=self._apply)
        apply_button.grid(row=0, column=1, padx=4)
        if not self.products_to_apply:
            apply_button.configure(state="disabled")

    def _apply(self) -> None:
        if not self.products_to_apply:
            return
        self.cancelled = False
        self.destroy()


class CostHistoryDialog(tk.Toplevel):
    def __init__(self, parent: OZPriceAnalyzerApp, rows: list[dict[str, object]]):
        super().__init__(parent)
        self.title("Журнал изменений себестоимости")
        self.geometry("1320x650")
        self.minsize(980, 500)
        self.transient(parent)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        ttk.Label(self, text="Журнал изменений себестоимости", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=20, pady=(18, 2)
        )
        ttk.Label(
            self,
            text="Журнал показывает ручные изменения, импорт XLSX и создание новых артикулов из отчетов.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 10))
        container = ttk.Frame(self)
        container.grid(row=2, column=0, sticky="nsew", padx=20)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        columns = ("date", "article", "name", "old_category", "new_category", "old_total", "new_total", "change", "old_labor", "new_labor", "source")
        tree = ttk.Treeview(container, columns=columns, show="headings")
        headings = [
            "Дата", "Артикул", "Наименование", "Старая категория", "Новая категория", "Старая с/с", "Новая с/с", "Изменение",
            "Старые трудозатраты", "Новые трудозатраты", "Источник",
        ]
        widths = [145, 130, 240, 180, 180, 120, 120, 120, 155, 155, 260]
        for column, heading, width in zip(columns, headings, widths):
            tree.heading(column, text=heading)
            tree.column(column, width=width, minwidth=80, stretch=False, anchor="w" if column in {"article", "name", "source"} else "e")
        xscroll = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        yscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        for row in rows:
            old_material = row["old_material_cost"]
            old_labor = row["old_labor_cost"]
            old_total = float(old_material) + float(old_labor) if old_material is not None and old_labor is not None else None
            new_total = float(row["new_material_cost"]) + float(row["new_labor_cost"])
            tree.insert(
                "",
                "end",
                values=(
                    str(row["changed_at"])[:16],
                    row["article"],
                    row["new_name"],
                    row["old_category"] or "—",
                    row["new_category"] or CATEGORY_EMPTY,
                    _money(old_total) if old_total is not None else "—",
                    _money(new_total),
                    _signed_money(new_total - old_total) if old_total is not None else "Новая",
                    _money(float(old_labor)) if old_labor is not None else "—",
                    _money(float(row["new_labor_cost"])),
                    row["change_source"],
                ),
            )
        if not rows:
            tree.insert("", "end", values=("", "", "Журнал пока пуст"))
        ttk.Button(self, text="Закрыть", command=self.destroy).grid(row=3, column=0, sticky="e", padx=20, pady=14)


class ProductDialog(tk.Toplevel):
    def __init__(self, parent: OZPriceAnalyzerApp, title: str, product: Product | None = None):
        super().__init__(parent)
        self.result: Product | None = None
        self.product = product
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.configure(background=parent.colors["window"])
        self.columnconfigure(1, weight=1)
        self.article_var = tk.StringVar(value=product.article if product else "")
        self.name_var = tk.StringVar(value=product.name if product else "")
        self.category_var = tk.StringVar(value=product.category if product else "")
        self.total_var = tk.StringVar(value=_plain_number(product.total_cost) if product else "")
        self.labor_var = tk.StringVar(value=_plain_number(product.labor_cost) if product else "0")

        fields = [
            ("Артикул", self.article_var),
            ("Наименование", self.name_var),
            ("Категория", self.category_var),
            ("Полная себестоимость, руб.", self.total_var),
            ("Трудозатраты в составе с/с, руб.", self.labor_var),
        ]
        for row, (label, variable) in enumerate(fields):
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=6, padx=(22, 12))
            entry = ttk.Entry(self, textvariable=variable, width=42)
            entry.grid(row=row, column=1, sticky="ew", pady=6, padx=(0, 22))
            if product and row == 0:
                entry.configure(state="disabled")
        buttons = ttk.Frame(self)
        buttons.grid(row=len(fields), column=0, columnspan=2, sticky="e", padx=18, pady=(14, 18))
        ttk.Button(buttons, text="Отмена", command=self.destroy).grid(row=0, column=0, padx=4)
        ttk.Button(buttons, text="Сохранить", style="Accent.TButton", command=self._save).grid(row=0, column=1, padx=4)
        self.bind("<Return>", lambda _event: self._save())
        self.bind("<Escape>", lambda _event: self.destroy())

    def _save(self) -> None:
        try:
            article = self.article_var.get().strip()
            name = self.name_var.get().strip() or article
            total = _parse_number(self.total_var.get())
            labor = _parse_number(self.labor_var.get())
            if not article or total < 0 or labor < 0 or labor > total:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Товар",
                "Укажите артикул и корректную себестоимость. Трудозатраты должны быть от 0 до полной себестоимости.",
                parent=self,
            )
            return
        self.result = Product(
            article=article,
            name=name,
            material_cost=total - labor,
            labor_cost=labor,
            active=self.product.active if self.product else True,
            category=self.category_var.get().strip(),
        )
        self.destroy()


class UnknownProductsDialog(tk.Toplevel):
    def __init__(self, parent: OZPriceAnalyzerApp, unknown: list[UnknownProduct]):
        super().__init__(parent)
        self.title("Новые товары")
        self.geometry("1040x620")
        self.transient(parent)
        self.grab_set()
        self.cancelled = True
        self.items = {item.article: item for item in unknown}
        self.decisions: dict[str, Product | None] = {}

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        ttk.Label(self, text="В отчетах найдены новые артикулы", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", padx=20, pady=(18, 2)
        )
        ttk.Label(
            self,
            text="Для включения начислений укажите себестоимость. Пропущенный артикул не попадет ни в товар, ни в нераспределенные суммы.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 10))

        container = ttk.Frame(self)
        container.grid(row=2, column=0, sticky="nsew", padx=20)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(container, columns=("article", "name", "sku", "sources", "decision"), show="headings")
        headings = ["Артикул", "Наименование", "SKU", "Файлы", "Решение"]
        widths = [150, 260, 120, 310, 130]
        for column, heading, width in zip(self.tree["columns"], headings, widths):
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, stretch=False, anchor="w")
        scroll = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        for item in unknown:
            self.tree.insert(
                "", "end", iid=item.article,
                values=(item.article, item.name, item.sku, ", ".join(sorted(item.source_names)), "Не выбрано"),
            )
        self.tree.bind("<<TreeviewSelect>>", self._select)

        editor = ttk.Frame(self, padding=(20, 12))
        editor.grid(row=3, column=0, sticky="ew")
        ttk.Label(editor, text="Полная себестоимость, руб.:").grid(row=0, column=0, padx=(0, 6))
        self.total_var = tk.StringVar()
        ttk.Entry(editor, textvariable=self.total_var, width=14).grid(row=0, column=1, padx=(0, 14))
        ttk.Label(editor, text="Трудозатраты, руб.:").grid(row=0, column=2, padx=(0, 6))
        self.labor_var = tk.StringVar(value="0")
        ttk.Entry(editor, textvariable=self.labor_var, width=14).grid(row=0, column=3, padx=(0, 14))
        ttk.Label(editor, text="Категория:").grid(row=1, column=0, padx=(0, 6), pady=(8, 0))
        self.category_var = tk.StringVar()
        ttk.Entry(editor, textvariable=self.category_var, width=42).grid(
            row=1, column=1, columnspan=3, sticky="ew", padx=(0, 14), pady=(8, 0)
        )
        ttk.Button(editor, text="Создать позицию", command=self._create).grid(row=0, column=4, padx=4)
        ttk.Button(editor, text="Пропустить", command=self._skip).grid(row=0, column=5, padx=4)

        buttons = ttk.Frame(self, padding=(20, 12))
        buttons.grid(row=4, column=0, sticky="e")
        ttk.Button(buttons, text="Отменить импорт", command=self.destroy).grid(row=0, column=0, padx=4)
        ttk.Button(buttons, text="Продолжить расчет", style="Accent.TButton", command=self._finish).grid(row=0, column=1, padx=4)
        first = next(iter(self.items), None)
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)

    @property
    def created_products(self) -> list[Product]:
        return [item for item in self.decisions.values() if item is not None]

    @property
    def skipped_articles(self) -> set[str]:
        return {article for article, product in self.decisions.items() if product is None}

    def _selected_article(self) -> str | None:
        selection = self.tree.selection()
        return selection[0] if selection else None

    def _select(self, _event=None) -> None:
        article = self._selected_article()
        decision = self.decisions.get(article) if article else None
        if isinstance(decision, Product):
            self.total_var.set(_plain_number(decision.total_cost))
            self.labor_var.set(_plain_number(decision.labor_cost))
            self.category_var.set(decision.category)
        else:
            self.total_var.set("")
            self.labor_var.set("0")
            self.category_var.set("")

    def _create(self) -> None:
        article = self._selected_article()
        if not article:
            return
        try:
            total = _parse_number(self.total_var.get())
            labor = _parse_number(self.labor_var.get())
            if total < 0 or labor < 0 or labor > total:
                raise ValueError
        except ValueError:
            messagebox.showerror("Себестоимость", "Проверьте полную себестоимость и трудозатраты", parent=self)
            return
        item = self.items[article]
        self.decisions[article] = Product(
            article,
            item.name or article,
            total - labor,
            labor,
            category=self.category_var.get().strip(),
        )
        self._set_decision_text(article, f"Создать: {_money(total)}")
        self._select_next_unresolved()

    def _skip(self) -> None:
        article = self._selected_article()
        if not article:
            return
        self.decisions[article] = None
        self._set_decision_text(article, "Пропустить")
        self._select_next_unresolved()

    def _set_decision_text(self, article: str, text: str) -> None:
        values = list(self.tree.item(article, "values"))
        values[-1] = text
        self.tree.item(article, values=values)

    def _select_next_unresolved(self) -> None:
        for article in self.items:
            if article not in self.decisions:
                self.tree.selection_set(article)
                self.tree.focus(article)
                self.tree.see(article)
                return

    def _finish(self) -> None:
        unresolved = [article for article in self.items if article not in self.decisions]
        if unresolved:
            messagebox.showwarning(
                "Новые товары",
                f"Выберите действие еще для {len(unresolved)} позиций: создать или пропустить.",
                parent=self,
            )
            return
        self.cancelled = False
        self.destroy()


def _result_values(result: ProductResult, tax_rate: float) -> tuple[object, ...]:
    return (
        result.article,
        result.name,
        result.category or CATEGORY_EMPTY,
        _money(result.total_cost),
        _money(result.material_cost),
        _money(result.labor_cost),
        _money(result.material_sold),
        _money(result.labor_sold),
        _money(result.cost_sold),
        _percent(result.profitability(tax_rate)),
        _money(result.net_profit_per_unit(tax_rate)),
        _money(result.profit_per_unit()),
        _money(result.net_profit(tax_rate)),
        _money(result.financial_result),
        _money(result.average_price()) if result.average_price() is not None else "—",
        _money(result.tax(tax_rate)),
        _money(result.taxable_income),
        _number(result.units),
        _money(result.revenue_including_points),
        _money(result.revenue_no_points),
        _money(result.partner_programs),
        _money(result.points),
        _money(result.commission),
        _money(result.processing),
        _money(result.delivery),
        _money(result.logistics),
        _money(result.reverse_logistics),
        _money(result.returns_cancels),
        _money(result.acquiring),
        _money(result.stars),
        _money(result.packaging),
        _money(result.compensation),
        _money(result.other),
        _money(result.financial_result),
        _percent(result.commission_share()),
        _percent(result.logistics_share()),
        _percent(result.points_share()),
        _percent(result.net_margin(tax_rate)),
    )


def _scenario_values(row: ScenarioRow) -> tuple[object, ...]:
    return (
        row.article,
        row.name,
        row.category or CATEGORY_EMPTY,
        _money(row.unit_cost),
        _number(row.units),
        _optional_money(row.current_price),
        _optional_money(row.planned_price),
        _optional_percent(row.price_change),
        _optional_percent(row.profitability),
        _optional_money(row.ozon_costs_without_commission),
        _optional_money(row.planned_revenue),
        _optional_percent(row.commission_rate),
        _optional_money(row.planned_commission),
        _optional_money(row.planned_points),
        _optional_money(row.taxable_base),
        _optional_money(row.tax),
        _optional_money(row.profit),
        _optional_money(row.profit_per_unit_before_cost),
        _optional_money(row.net_profit_per_unit),
        _optional_money(row.net_profit_total),
    )


def filter_product_results(
    rows: list[ProductResult],
    tax_rate: float,
    *,
    category: str = CATEGORY_ALL,
    article_query: str = "",
    sort_metric: str = SORT_NONE,
    descending: bool = False,
) -> list[ProductResult]:
    visible = _filter_rows(rows, category, article_query)
    metrics = {
        "Доходность": lambda row: row.profitability(tax_rate),
        "Чистая прибыль": lambda row: row.net_profit(tax_rate),
        "Выручка": lambda row: row.revenue_including_points,
        "Количество продаж": lambda row: row.units,
    }
    return _sort_rows(visible, metrics.get(sort_metric), descending)


def summarize_category(
    rows: list[ProductResult],
    tax_rate: float,
    category: str = CATEGORY_ALL,
) -> dict[str, float | int]:
    """Aggregate product-only KPIs for one category; unallocated rows are not inputs."""
    selected = [
        row
        for row in rows
        if category == CATEGORY_ALL or _category_label(row.category) == category
    ]
    revenue = sum(row.revenue_including_points for row in selected)
    net_profit = sum(row.net_profit(tax_rate) for row in selected)
    cost_sold = sum(row.cost_sold for row in selected)
    return {
        "product_count": len(selected),
        "units": sum(row.units for row in selected),
        "revenue": revenue,
        "cost_sold": cost_sold,
        "tax": sum(row.tax(tax_rate) for row in selected),
        "financial_result": sum(row.financial_result for row in selected),
        "net_profit": net_profit,
        "profitability": net_profit / cost_sold if cost_sold else 0.0,
    }


def filter_scenario_rows(
    rows: list[ScenarioRow],
    *,
    category: str = CATEGORY_ALL,
    article_query: str = "",
    sort_metric: str = SORT_NONE,
    descending: bool = False,
) -> list[ScenarioRow]:
    visible = _filter_rows(rows, category, article_query)
    metrics = {
        "Доходность": lambda row: row.profitability,
        "Чистая прибыль": lambda row: row.net_profit_total,
        "Выручка": lambda row: row.planned_revenue,
        "Количество продаж": lambda row: row.units,
    }
    return _sort_rows(visible, metrics.get(sort_metric), descending)


def _filter_rows(rows, category: str, article_query: str):
    query = article_query.strip().casefold()
    return [
        row
        for row in rows
        if (category == CATEGORY_ALL or _category_label(row.category) == category)
        and (not query or query in row.article.casefold())
    ]


def _sort_rows(rows, metric, descending: bool):
    if metric is None:
        return list(rows)
    measured: list[tuple[float, object]] = []
    missing: list[object] = []
    for row in rows:
        value = metric(row)
        if value is None:
            missing.append(row)
        else:
            measured.append((float(value), row))
    measured.sort(key=lambda pair: (pair[0], pair[1].article.casefold()), reverse=descending)
    return [row for _value, row in measured] + missing


def _category_label(value: str) -> str:
    return value.strip() or CATEGORY_EMPTY


def _russian_position_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "позиция"
    if count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        return "позиции"
    return "позиций"


def _russian_report_count(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        word = "отчет"
    elif count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        word = "отчета"
    else:
        word = "отчетов"
    return f"{count} {word}"


def _summary_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def _run_year(run: RunSummary) -> int | None:
    value = (
        _summary_date(run.period_start)
        or _summary_date(run.period_end)
        or _summary_date(run.created_at)
    )
    return value.year if value is not None else None


def _realization_file_count(calculation: RunCalculation) -> str:
    count = sum(
        source.report_type == REPORT_REALIZATION
        for source in calculation.source_files
    )
    if count % 10 == 1 and count % 100 != 11:
        word = "файл выкупов"
    elif count % 10 in {2, 3, 4} and count % 100 not in {12, 13, 14}:
        word = "файла выкупов"
    else:
        word = "файлов выкупов"
    return f"{count} {word}"


def _set_category_choices(combo: ttk.Combobox, variable: tk.StringVar, categories) -> None:
    values = [CATEGORY_ALL] + sorted(
        {_category_label(str(value or "")) for value in categories},
        key=str.casefold,
    )
    combo["values"] = values
    if variable.get() not in values:
        variable.set(CATEGORY_ALL)


def _money(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value:,.2f} ₽".replace(",", " ")


def _signed_money(value: float) -> str:
    return ("+" if value > 0 else "") + _money(value)


def _signed_number(value: float) -> str:
    return ("+" if value > 0 else "") + _number(value)


def _number(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ").rstrip("0").rstrip(".")


def _percent(value: float) -> str:
    return f"{value * 100:,.2f}%".replace(",", " ")


def _signed_percentage_points(value: float) -> str:
    prefix = "+" if value > 0 else ""
    return f"{prefix}{value * 100:,.2f} п.п.".replace(",", " ")


def _comparison_percent(metric: ComparisonMetric) -> str:
    value = metric.change_percent
    if value is None:
        return "0,00%"
    if value == float("inf"):
        return "новое значение"
    prefix = "+" if value > 0 else ""
    return prefix + _percent(value)


def _comparison_kpi(metric: ComparisonMetric, money: bool) -> str:
    absolute = _signed_money(metric.change) if money else _signed_number(metric.change)
    return f"{absolute} · {_comparison_percent(metric)}"


def _axis_value(value: float, metric: str) -> str:
    if metric == "units":
        return _number(value)
    if metric in TREND_PERCENT_METRICS:
        return _percent(value)
    absolute = abs(value)
    if absolute >= 1_000_000:
        return f"{value / 1_000_000:.1f} млн"
    if absolute >= 1_000:
        return f"{value / 1_000:.0f} тыс."
    return f"{value:.0f}"


def _trend_value(value: float, metric: str) -> str:
    if metric == "units":
        return _number(value)
    if metric in TREND_PERCENT_METRICS:
        return _percent(value)
    return _money(value)


def _short_period(value: str) -> str:
    return value.split("–", 1)[0]


def _optional_money(value: float | None) -> str:
    return _money(value) if value is not None else "—"


def _optional_percent(value: float | None) -> str:
    return _percent(value) if value is not None else "—"


def _plain_number(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.6f}".rstrip("0").rstrip(".")


def _parse_number(value: str) -> float:
    return float(value.replace("\u00a0", "").replace(" ", "").replace(",", ".").replace("₽", "").replace("%", "").strip())


def _period_text(start: str | None, end: str | None) -> str:
    if start and end:
        return f"{_date_display(start)}–{_date_display(end)}"
    return "Период не определен"


def _date_display(value: str) -> str:
    parts = value[:10].split("-")
    return ".".join(reversed(parts)) if len(parts) == 3 else value


def _backup_timestamp(value: str) -> str:
    try:
        return datetime.fromisoformat(value).strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return value or "дата не указана"


def _file_size(value: int) -> str:
    size = float(value)
    for unit in ("Б", "КБ", "МБ", "ГБ"):
        if size < 1024 or unit == "ГБ":
            return f"{size:.0f} {unit}" if unit == "Б" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{value} Б"


def _calculation_period(calculation: RunCalculation) -> str:
    if calculation.period_start and calculation.period_end:
        return f"{calculation.period_start:%d.%m.%Y}–{calculation.period_end:%d.%m.%Y}"
    return "не определен"


def _session_period(session: ImportSession) -> str:
    if session.period_start and session.period_end:
        return f"{session.period_start:%d.%m.%Y}–{session.period_end:%d.%m.%Y}"
    return "не определен"


def _run_positions(runs) -> dict[int, int]:
    return {run.id: position for position, run in enumerate(runs, start=1)}


def _run_years(run) -> set[int]:
    start_year = _text_year(run.period_start)
    end_year = _text_year(run.period_end)
    if start_year is not None or end_year is not None:
        first = start_year if start_year is not None else end_year
        last = end_year if end_year is not None else start_year
        assert first is not None and last is not None
        lower, upper = sorted((first, last))
        return set(range(lower, upper + 1))
    created_year = _text_year(run.created_at)
    return {created_year} if created_year is not None else set()


def _filter_runs_by_years(runs, selected_years: set[int] | None):
    if selected_years is None:
        return list(runs)
    return [run for run in runs if _run_years(run) & selected_years]


def _year_filter_label(selected_years: set[int] | None) -> str:
    if selected_years is None:
        return "Все годы"
    years = sorted(selected_years)
    if len(years) <= 3:
        return ", ".join(str(year) for year in years)
    return f"Выбрано лет: {len(years)}"


def _current_year_period_label() -> str:
    return f"Текущий год ({date.today().year})"


def _trend_period_label(mode: str, selected_years: set[int] | None) -> str:
    if mode == "current":
        return _current_year_period_label()
    return _year_filter_label(selected_years)


def _text_year(value: str | None) -> int | None:
    if not value:
        return None
    try:
        return int(value[:4])
    except (TypeError, ValueError):
        return None


def _open_path(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def run_app() -> None:
    from .single_instance import launch_single_instance

    launch_single_instance(OZPriceAnalyzerApp)


if __name__ == "__main__":
    run_app()
