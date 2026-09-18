from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .overview_column_settings import OverviewColumnSettingsOZPriceAnalyzerApp
from .ui import _money, _percent


def report_total_value(calculation) -> float:
    """Return report profit after tax on positive unallocated Ozon income."""
    return float(calculation.report_net_profit)


def report_total_profitability(calculation) -> float:
    """Return report profit after tax divided by the cost of goods sold."""
    cost_sold = float(calculation.totals()["cost_sold"])
    return report_total_value(calculation) / cost_sold if cost_sold else 0.0


def overview_revenue_kpi_values(calculation) -> dict[str, str]:
    """Format paired monetary and relative KPIs for the Overview tab."""
    amounts = calculation.revenue_amounts()
    shares = calculation.revenue_shares()
    return {
        "commission": f"{_money(amounts['commission'])} · {_percent(shares['commission_share'])}",
        "logistics": f"{_money(amounts['logistics'])} · {_percent(shares['logistics_share'])}",
        "points": f"{_money(amounts['points'])} · {_percent(shares['points_share'])}",
        "net_margin": _percent(shares["net_margin"]),
    }


class ReportTotalsOZPriceAnalyzerApp(OverviewColumnSettingsOZPriceAnalyzerApp):
    """Report totals, revenue shares and configurable report-table columns."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._install_report_total_kpi()
        self._install_revenue_share_kpis()
        self._clarify_financial_headings()
        self.after_idle(self._refresh_report_total_kpi)
        self.after_idle(self._refresh_revenue_share_kpis)

    def _install_report_total_kpi(self) -> None:
        self.kpi_frame.columnconfigure(6, weight=1)
        for child in self.kpi_frame.winfo_children():
            try:
                info = child.grid_info()
                if int(info.get("row", -1)) in (0, 2):
                    child.grid_configure(columnspan=7)
            except (TypeError, ValueError):
                continue

        self.kpi_vars["report_total"] = self.kpi_vars.get("report_total") or tk.StringVar(
            master=self,
            value="—",
        )
        self.kpi_vars["report_total_profitability"] = self.kpi_vars.get(
            "report_total_profitability"
        ) or tk.StringVar(master=self, value="—")
        card = ttk.Frame(self.kpi_frame, style="Card.TFrame", padding=(8, 5))
        card.grid(row=1, column=6, sticky="nsew", padx=(5, 0))
        ttk.Label(
            card,
            text="Итог с нераспределёнными после налога",
            style="CompactCardMuted.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(
            card,
            textvariable=self.kpi_vars["report_total"],
            style="CompactKpi.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(1, 0))
        ttk.Label(
            card,
            text="Доходность:",
            style="CompactCardMuted.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(2, 0))
        ttk.Label(
            card,
            textvariable=self.kpi_vars["report_total_profitability"],
            style="CompactCard.TLabel",
        ).grid(row=2, column=1, sticky="w", padx=(5, 0), pady=(2, 0))

    def _install_revenue_share_kpis(self) -> None:
        for child in self.kpi_frame.winfo_children():
            try:
                row = int(child.grid_info().get("row", -1))
            except (TypeError, ValueError):
                continue
            if row == 2:
                child.grid_configure(row=3)
            elif row == 3:
                child.grid_configure(row=4)

        share_frame = ttk.Frame(self.kpi_frame)
        share_frame.grid(row=2, column=0, columnspan=7, sticky="ew", pady=(3, 2))
        for column in range(4):
            share_frame.columnconfigure(column, weight=1)

        self.revenue_share_kpi_vars = {}
        cards = (
            ("commission", "Комиссия Ozon"),
            ("logistics", "Логистика"),
            ("points", "Баллы"),
            ("net_margin", "Чистая прибыль, %"),
        )
        for index, (key, title) in enumerate(cards):
            variable = tk.StringVar(master=self, value="—")
            self.revenue_share_kpi_vars[key] = variable
            card = ttk.Frame(share_frame, style="Card.TFrame", padding=(8, 5))
            card.columnconfigure(1, weight=1)
            card.grid(
                row=0,
                column=index,
                sticky="nsew",
                padx=(0 if index == 0 else 5, 0 if index == len(cards) - 1 else 5),
            )
            ttk.Label(card, text=title, style="CompactCardMuted.TLabel").grid(
                row=0, column=0, sticky="w"
            )
            ttk.Label(card, textvariable=variable, style="CompactKpi.TLabel").grid(
                row=0, column=1, sticky="e", padx=(8, 0)
            )

    def _clarify_financial_headings(self) -> None:
        self.overview_tree.heading("profit_unit", text="Финрезультат Ozon на ед.")
        self.overview_tree.heading(
            "profit_total",
            text="Финрезультат Ozon до с/с и налога",
        )
        self.scenario_tree.heading("profit", text="Прибыль до себестоимости")
        if hasattr(self, "_heading_tooltips"):
            self.after_idle(self._heading_tooltips.fit_all)

    def _populate_overview(self) -> None:
        super()._populate_overview()
        self._refresh_report_total_kpi()
        self._refresh_revenue_share_kpis()

    def _refresh_report_total_kpi(self) -> None:
        variable = self.kpi_vars.get("report_total")
        profitability_variable = self.kpi_vars.get("report_total_profitability")
        if variable is None or profitability_variable is None:
            return
        calculation = self.overview_calculation
        if calculation is None:
            variable.set("—")
            profitability_variable.set("—")
            return
        variable.set(_money(report_total_value(calculation)))
        profitability_variable.set(_percent(report_total_profitability(calculation)))

    def _refresh_revenue_share_kpis(self) -> None:
        variables = getattr(self, "revenue_share_kpi_vars", None)
        if not variables:
            return
        calculation = self.overview_calculation
        if calculation is None:
            for variable in variables.values():
                variable.set("—")
            return
        values = overview_revenue_kpi_values(calculation)
        for key, variable in variables.items():
            variable.set(values[key])


def run_app() -> None:
    from .single_instance import launch_single_instance

    launch_single_instance(ReportTotalsOZPriceAnalyzerApp)
