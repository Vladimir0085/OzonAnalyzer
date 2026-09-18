from __future__ import annotations

from tkinter import filedialog, messagebox, ttk

from .exporter import export_calculation, suggested_export_name
from .history_batch_export import HistoryBatchExportOZPriceAnalyzerApp


class OverviewExportOZPriceAnalyzerApp(HistoryBatchExportOZPriceAnalyzerApp):
    """v0.5.15: export the calculation currently shown on the Overview tab."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._install_overview_export_button()

    def _install_overview_export_button(self) -> None:
        overview_header = getattr(self, "overview_header", None)
        if overview_header is None:
            return

        self.overview_export_button = ttk.Button(
            overview_header,
            text="Выгрузить итог XLSX",
            style="Compact.TButton",
            command=self.export_overview_calculation,
        )
        self.overview_export_button.grid(row=0, column=4, padx=(8, 0))

    def export_overview_calculation(self) -> None:
        calculation = self.overview_calculation
        if calculation is None:
            messagebox.showinfo(
                "Экспорт обзора",
                "Сначала выберите один или несколько отчетов для обзора.",
                parent=self,
            )
            return

        destination = filedialog.asksaveasfilename(
            title="Сохранить итоговый отчет обзора",
            defaultextension=".xlsx",
            initialdir=str(self.service.paths["exports"]),
            initialfile=suggested_export_name(calculation),
            filetypes=[("Книга Excel", "*.xlsx")],
            parent=self,
        )
        if not destination:
            return

        try:
            path = export_calculation(calculation, destination)
            report_count = len(self.overview_run_ids)
            description = (
                f"Объединенный отчет по {report_count} отчетам"
                if report_count > 1
                else "Итоговый отчет"
            )
            messagebox.showinfo(
                "Экспорт обзора",
                f"{description} сохранен:\n{path}",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Экспорт обзора", str(exc), parent=self)


def run_app() -> None:
    app = OverviewExportOZPriceAnalyzerApp()
    app.mainloop()
