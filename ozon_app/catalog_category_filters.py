from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Iterable, TypeVar

from .calculator import calculate_scenario
from .display_modes import LaptopFriendlyOZPriceAnalyzerApp, _button_with_text
from .ui import (
    CATEGORY_ALL,
    CATEGORY_EMPTY,
    SORT_ASCENDING,
    SORT_DESCENDING,
    SORT_NONE,
    _category_label,
    _money,
    _number,
    _profitability_text,
    _result_values,
    _russian_position_word,
    _russian_report_count,
    _scenario_values,
    category_profitability_text,
    filter_product_results,
    filter_scenario_rows,
    scenario_tax_note,
    summarize_category,
)


T = TypeVar("T")
FILTER_INCLUDE = "include"
FILTER_EXCLUDE = "exclude"


def apply_category_filter(
    rows: Iterable[T],
    selected_categories: set[str] | frozenset[str],
    *,
    exclude: bool = False,
) -> list[T]:
    """Filter rows by their ``category`` attribute.

    Empty selection intentionally means "all categories" in either mode.
    """
    selected = {str(value) for value in selected_categories if str(value).strip()}
    if not selected:
        return list(rows)

    result: list[T] = []
    for row in rows:
        label = _category_label(str(getattr(row, "category", "") or ""))
        matches = label in selected
        if (exclude and not matches) or (not exclude and matches):
            result.append(row)
    return result


def category_filter_button_text(selected_categories: set[str] | frozenset[str]) -> str:
    selected = sorted(selected_categories, key=str.casefold)
    if not selected:
        return "Все категории"
    if len(selected) == 1:
        return selected[0]
    if len(selected) == 2:
        return " · ".join(selected)
    return f"Выбрано категорий: {len(selected)}"


def category_filter_scope_text(
    selected_categories: set[str] | frozenset[str],
    *,
    exclude: bool = False,
) -> str:
    selected = sorted(selected_categories, key=str.casefold)
    if not selected:
        return "Все товары"
    names = ", ".join(selected[:3])
    if len(selected) > 3:
        names += f" и ещё {len(selected) - 3}"
    return f"кроме категорий: {names}" if exclude else f"категории: {names}"


def delete_catalog_product(database, article: str) -> bool:
    """Delete one product from the current cost catalog only.

    Saved report snapshots and scenario history are intentionally untouched.
    """
    normalized = str(article or "").strip()
    if not normalized:
        return False
    with database.transaction() as db:
        cursor = db.execute("DELETE FROM products WHERE article = ?", (normalized,))
        return bool(cursor.rowcount)


class CategoryMultiSelectDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        title: str,
        categories: list[str],
        selected_categories: set[str],
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.geometry("540x540")
        self.minsize(440, 420)
        self.transient(parent)
        self.grab_set()
        self.confirmed = False
        self.categories = list(categories)
        self.selected_categories = set(selected_categories) & set(categories)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        ttk.Label(
            self,
            text="Выберите одну или несколько категорий",
            style="Section.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(18, 3))
        ttk.Label(
            self,
            text=(
                "Обычный щелчок включает или выключает категорию. "
                "Пустой выбор означает «Все категории»."
            ),
            style="Muted.TLabel",
            wraplength=500,
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 10))

        container = ttk.Frame(self)
        container.grid(row=2, column=0, sticky="nsew", padx=20)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(
            container,
            selectmode=tk.MULTIPLE,
            exportselection=False,
            activestyle="dotbox",
            borderwidth=1,
            highlightthickness=0,
        )
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        for index, category in enumerate(self.categories):
            self.listbox.insert("end", category)
            if category in self.selected_categories:
                self.listbox.selection_set(index)

        controls = ttk.Frame(self)
        controls.grid(row=3, column=0, sticky="ew", padx=20, pady=(10, 0))
        controls.columnconfigure(2, weight=1)
        ttk.Button(controls, text="Выбрать все", command=self._select_all).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(controls, text="Снять все", command=self._clear).grid(
            row=0, column=1, padx=(0, 6)
        )
        self.summary_var = tk.StringVar()
        ttk.Label(controls, textvariable=self.summary_var, style="Muted.TLabel").grid(
            row=0, column=2, sticky="e"
        )

        footer = ttk.Frame(self, padding=(20, 14, 20, 18))
        footer.grid(row=4, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Button(footer, text="Отмена", command=self.destroy).grid(
            row=0, column=1, padx=4
        )
        ttk.Button(
            footer,
            text="Применить",
            style="Accent.TButton",
            command=self._confirm,
        ).grid(row=0, column=2, padx=4)

        self.listbox.bind("<<ListboxSelect>>", lambda _event: self._update_summary())
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Return>", lambda _event: self._confirm())
        self._update_summary()

    def _select_all(self) -> None:
        if self.categories:
            self.listbox.selection_set(0, "end")
        self._update_summary()

    def _clear(self) -> None:
        self.listbox.selection_clear(0, "end")
        self._update_summary()

    def _update_summary(self) -> None:
        count = len(self.listbox.curselection())
        self.summary_var.set("Все категории" if count == 0 else f"Выбрано: {count}")

    def _confirm(self) -> None:
        self.selected_categories = {
            self.categories[int(index)] for index in self.listbox.curselection()
        }
        self.confirmed = True
        self.destroy()


class CatalogAndCategoryOZPriceAnalyzerApp(LaptopFriendlyOZPriceAnalyzerApp):
    """v0.5.5: product deletion and multi-category include/exclude filters."""

    def __init__(self, *args, **kwargs) -> None:
        self.overview_selected_categories: set[str] = set()
        self.scenario_selected_categories: set[str] = set()
        self.overview_available_categories: list[str] = []
        self.scenario_available_categories: list[str] = []
        super().__init__(*args, **kwargs)

    def _build_overview_tab(self) -> None:
        super()._build_overview_tab()
        self._install_multi_category_control("overview")

    def _build_scenario_tab(self) -> None:
        super()._build_scenario_tab()
        self._install_multi_category_control("scenario")

    def _build_settings_tab(self) -> None:
        super()._build_settings_tab()
        journal_button = _button_with_text(self.settings_tab, "Журнал изменений")
        if journal_button is not None:
            header = journal_button.master
            ttk.Button(
                header,
                text="Удалить",
                style="Danger.TButton",
                command=self.delete_selected_product,
            ).grid(row=0, column=6, padx=4)
            for child in header.winfo_children():
                if isinstance(child, ttk.Label):
                    try:
                        if "При первом запуске справочник пуст" in str(child.cget("text")):
                            child.grid_configure(columnspan=7)
                    except tk.TclError:
                        pass
        self.products_tree.bind("<Delete>", lambda _event: self.delete_selected_product(), add="+")

    def _install_multi_category_control(self, key: str) -> None:
        combo = getattr(self, f"{key}_category_combo")
        info = dict(combo.grid_info())
        combo.grid_remove()
        parent = combo.master

        selected = (
            self.overview_selected_categories
            if key == "overview"
            else self.scenario_selected_categories
        )
        label_var = tk.StringVar(value=category_filter_button_text(selected))
        exclude_var = tk.BooleanVar(value=False)
        setattr(self, f"{key}_category_filter_label_var", label_var)
        setattr(self, f"{key}_category_exclude_var", exclude_var)

        frame = ttk.Frame(parent)
        frame.grid(
            row=int(info.get("row", 0)),
            column=int(info.get("column", 1)),
            sticky="w",
            padx=info.get("padx", (0, 14)),
            pady=info.get("pady", 0),
        )
        setattr(self, f"{key}_category_filter_frame", frame)

        ttk.Button(
            frame,
            textvariable=label_var,
            command=lambda k=key: self._choose_categories(k),
            width=21,
            style="Compact.TButton" if key == "overview" else "TButton",
        ).grid(row=0, column=0, padx=(0, 6))
        ttk.Checkbutton(
            frame,
            text="Исключить выбранные",
            variable=exclude_var,
            command=lambda k=key: self._on_category_mode_changed(k),
        ).grid(row=0, column=1)

    def _available_category_labels(self, values: Iterable[str]) -> list[str]:
        return sorted(
            {_category_label(str(value or "")) for value in values},
            key=str.casefold,
        )

    def _sync_filter_categories(self, key: str, values: Iterable[str]) -> None:
        available = self._available_category_labels(values)
        setattr(self, f"{key}_available_categories", available)
        selected: set[str] = getattr(self, f"{key}_selected_categories")
        selected.intersection_update(available)
        self._update_category_filter_label(key)

    def _update_category_filter_label(self, key: str) -> None:
        variable = getattr(self, f"{key}_category_filter_label_var", None)
        if variable is None:
            return
        selected: set[str] = getattr(self, f"{key}_selected_categories")
        variable.set(category_filter_button_text(selected))

    def _choose_categories(self, key: str) -> None:
        available: list[str] = getattr(self, f"{key}_available_categories")
        if not available:
            messagebox.showinfo(
                "Фильтр категорий",
                "В текущем отчете нет категорий для выбора.",
                parent=self,
            )
            return
        selected: set[str] = getattr(self, f"{key}_selected_categories")
        dialog = CategoryMultiSelectDialog(
            self,
            title="Категории — Обзор" if key == "overview" else "Категории — Сценарий цены",
            categories=available,
            selected_categories=selected,
        )
        self.wait_window(dialog)
        if not dialog.confirmed:
            return
        selected.clear()
        selected.update(dialog.selected_categories)
        self._update_category_filter_label(key)
        if key == "overview":
            self._populate_overview()
        else:
            self._populate_scenario()

    def _on_category_mode_changed(self, key: str) -> None:
        if key == "overview":
            self._populate_overview()
        else:
            self._populate_scenario()

    def _filter_scope(self, key: str) -> tuple[set[str], bool]:
        selected: set[str] = getattr(self, f"{key}_selected_categories")
        variable = getattr(self, f"{key}_category_exclude_var", None)
        exclude = bool(variable.get()) if variable is not None else False
        return selected, exclude

    def _populate_overview(self) -> None:
        calculation = self.overview_calculation
        if calculation is None:
            return

        totals = calculation.totals()
        cost = totals["cost_sold"]
        self.kpi_vars["revenue"].set(_money(totals["revenue"]))
        self.kpi_vars["net_profit"].set(_money(totals["net_profit"]))
        self.kpi_vars["profitability"].set(
            _profitability_text(
                totals["net_profit"] / cost if cost else 0.0,
                units=totals["units"],
                cost_sold=cost,
            )
        )
        self.kpi_vars["units"].set(_number(totals["units"]))
        self.kpi_vars["unallocated"].set(_money(totals["unallocated"]))
        self.kpi_vars["files"].set(str(self.overview_file_count))

        self._sync_filter_categories(
            "overview", (result.category for result in calculation.products)
        )
        selected_categories, exclude = self._filter_scope("overview")
        category_rows = apply_category_filter(
            calculation.products,
            selected_categories,
            exclude=exclude,
        )

        category_totals = summarize_category(category_rows, calculation.tax_rate, CATEGORY_ALL)
        scope = category_filter_scope_text(selected_categories, exclude=exclude)
        position_word = _russian_position_word(int(category_totals["product_count"]))
        self.category_summary_title_var.set(
            f"Итоги по фильтру: {scope} · "
            f"{int(category_totals['product_count'])} {position_word}"
        )
        self.category_kpi_vars["revenue"].set(_money(category_totals["revenue"]))
        self.category_kpi_vars["net_profit"].set(_money(category_totals["net_profit"]))
        self.category_kpi_vars["profitability"].set(category_profitability_text(category_totals))
        self.category_kpi_vars["units"].set(_number(category_totals["units"]))
        self.category_kpi_vars["cost_sold"].set(_money(category_totals["cost_sold"]))
        self.category_kpi_vars["financial_result"].set(_money(category_totals["financial_result"]))

        visible = filter_product_results(
            category_rows,
            calculation.tax_rate,
            category=CATEGORY_ALL,
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

    def _reset_overview_filters(self) -> None:
        self.overview_selected_categories.clear()
        if hasattr(self, "overview_category_exclude_var"):
            self.overview_category_exclude_var.set(False)
        self._update_category_filter_label("overview")
        self.overview_article_var.set("")
        self.overview_sort_var.set(SORT_NONE)
        self.overview_sort_direction_var.set(SORT_ASCENDING)
        self._populate_overview()

    def _populate_scenario(self) -> None:
        calculation = self._active_report_calculation()
        self.scenario_tree.delete(*self.scenario_tree.get_children())
        self.scenario_rows.clear()
        if calculation is None:
            for variable in self.scenario_kpi_vars.values():
                variable.set("—")
            self.scenario_tax_note_var.set("")
            self.scenario_count_var.set("")
            return

        target_run_id = self._active_single_run_id()
        prices = self.db.planned_prices(target_run_id) if target_run_id is not None else {}
        planned_revenue_total = 0.0
        planned_net_total = 0.0
        planned_cost_total = 0.0
        scenarios = []

        for result in calculation.products:
            scenario = calculate_scenario(result, calculation.tax_rate, prices.get(result.article))
            scenarios.append(scenario)
            self.scenario_rows[result.article] = scenario
            if scenario.planned_revenue is not None:
                planned_revenue_total += scenario.planned_revenue
            if scenario.net_profit_per_unit is not None:
                planned_net_total += scenario.net_profit_per_unit * scenario.units
                planned_cost_total += scenario.unit_cost * scenario.units

        self._sync_filter_categories("scenario", (row.category for row in scenarios))
        selected_categories, exclude = self._filter_scope("scenario")
        category_rows = apply_category_filter(
            scenarios,
            selected_categories,
            exclude=exclude,
        )
        visible = filter_scenario_rows(
            category_rows,
            category=CATEGORY_ALL,
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
        self.scenario_kpi_vars["planned_margin"].set(
            _profitability_text(
                planned_net_total / planned_cost_total if planned_cost_total else 0.0,
                units=totals["units"],
                cost_sold=planned_cost_total,
            )
        )
        self.scenario_tax_note_var.set(scenario_tax_note(calculation))
        scope_note = ""
        active_count = len(self._active_report_run_ids())
        if active_count > 1:
            scope_note = f" · {_russian_report_count(active_count)} · цены только для 1 отчета"
        self.scenario_count_var.set(
            f"Показано: {len(visible)} из {len(scenarios)}{scope_note}"
        )
        self._configure_value_tags(self.scenario_tree)

    def _reset_scenario_filters(self) -> None:
        self.scenario_selected_categories.clear()
        if hasattr(self, "scenario_category_exclude_var"):
            self.scenario_category_exclude_var.set(False)
        self._update_category_filter_label("scenario")
        self.scenario_article_var.set("")
        self.scenario_sort_var.set(SORT_NONE)
        self.scenario_sort_direction_var.set(SORT_ASCENDING)
        self._populate_scenario()

    def delete_selected_product(self) -> None:
        selection = self.products_tree.selection()
        if not selection:
            messagebox.showinfo(
                "Удалить товар",
                "Выберите позицию в справочнике себестоимости.",
                parent=self,
            )
            return

        article = selection[0]
        product = self.db.product_map(active_only=False).get(article)
        if product is None:
            self.refresh_products()
            return

        if not messagebox.askyesno(
            "Удалить позицию из справочника?",
            f"Удалить «{product.name}» (артикул {product.article}) "
            "из текущего справочника себестоимости?\n\n"
            "Сохраненные отчеты, их показатели и историческая себестоимость "
            "останутся без изменений. Если этот артикул встретится в новом отчете, "
            "его потребуется добавить заново.",
            icon="warning",
            parent=self,
        ):
            return

        if not delete_catalog_product(self.db, article):
            messagebox.showinfo(
                "Удалить товар",
                "Позиция уже отсутствует в текущем справочнике.",
                parent=self,
            )
            self.refresh_products()
            return

        self._refresh_after_catalog_change()
        self.status_var.set(f"Из справочника удалена позиция: {article}")


def run_app() -> None:
    app = CatalogAndCategoryOZPriceAnalyzerApp()
    app.mainloop()
