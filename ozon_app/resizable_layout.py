from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .ui import OZPriceAnalyzerApp


def _minimum_upper_height(configured: int, protected_heights) -> int:
    """Return a safe upper-pane minimum that cannot collapse protected rows."""
    heights = [max(0, int(value)) for value in protected_heights]
    protected = sum(heights)
    return max(int(configured), protected + (12 if protected else 0))


class _GridSplitter:
    """Static table toolbar that keeps controls separate from the table.

    The component intentionally preserves the historical API used by display
    modes, but vertical dragging is disabled. By default the table expands into
    the remaining space; an optional fraction keeps Overview panes proportional
    at every window size.
    """

    def __init__(self, owner, parent, *, row: int, table_row: int, absorb_row: int,
                 min_upper: int, min_table: int, compact=None, protected_rows=(),
                 table_fraction: float | None = None, bar: ttk.Frame | None = None):
        self.owner = owner
        self.parent = parent
        self.row = row
        self.table_row = table_row
        self.absorb_row = absorb_row
        self.min_upper = min_upper
        self.min_table = min_table
        self.compact = compact
        self.protected_rows = tuple(int(value) for value in protected_rows)
        self.table_fraction = (
            None
            if table_fraction is None
            else max(0.4, min(float(table_fraction), 0.7))
        )
        self.dragging = False
        self.suspended = False
        self.table_height: int | None = None
        self._control_column = 1

        self.bar = bar if bar is not None else ttk.Frame(parent)
        self.control_row = 1 if bar is not None else 0
        self.bar.grid(row=row, column=0, sticky="ew", pady=2)
        self.bar.columnconfigure(0, weight=1)

        # This is now a visual separator only. It is deliberately not bound to
        # mouse events and does not show a resize cursor.
        self.handle = ttk.Separator(self.bar, orient="horizontal")
        self.handle.grid(row=self.control_row, column=0, sticky="ew", padx=(0, 6))

        parent.bind("<Configure>", self._on_configure, add="+")
        owner.after_idle(self._initialize)

    def add_button(self, text: str, command) -> ttk.Button:
        button = ttk.Button(
            self.bar,
            text=text,
            style="TableTool.TButton",
            command=command,
        )
        button.grid(row=self.control_row, column=self._control_column, padx=(4, 0))
        self._control_column += 1
        return button

    def _protected_row_heights(self) -> list[int]:
        if not self.protected_rows:
            return []
        heights = {row: 0 for row in self.protected_rows}
        for child in self.parent.winfo_children():
            if child.winfo_manager() != "grid":
                continue
            try:
                row = int(child.grid_info().get("row", -1))
            except (TypeError, ValueError, tk.TclError):
                continue
            if row in heights:
                try:
                    heights[row] = max(heights[row], int(child.winfo_reqheight()))
                except tk.TclError:
                    pass
        return [heights[row] for row in self.protected_rows]

    def _effective_min_upper(self) -> int:
        return _minimum_upper_height(self.min_upper, self._protected_row_heights())

    def _initialize(self):
        if self.suspended:
            return
        total = self.parent.winfo_height()
        if total <= 1:
            self.owner.after(40, self._initialize)
            return
        self._apply()

    # Kept as no-ops for backwards compatibility with code that may still call
    # these methods. Table height is no longer user-draggable.
    def _start(self, _event):
        self.dragging = False

    def _move(self, _event):
        return

    def _finish(self, _event):
        self.dragging = False

    def _on_configure(self, _event):
        if not self.suspended:
            self.owner.after_idle(self._apply)

    def _apply(self):
        if self.suspended:
            return
        total = self.parent.winfo_height()
        if total <= 1:
            return

        if self.table_fraction is not None:
            # Overview uses the same adaptive split as the WB application: the
            # compact controls and the table receive equal shares of the usable
            # area at every window and monitor size.
            group = f"split_{id(self.parent)}"
            upper_weight = max(1, round((1.0 - self.table_fraction) * 100))
            table_weight = max(1, round(self.table_fraction * 100))
            self.table_height = None
            self.parent.rowconfigure(
                self.absorb_row,
                weight=upper_weight,
                minsize=0,
                uniform=group,
            )
            self.parent.rowconfigure(
                self.table_row,
                weight=table_weight,
                minsize=max(110, min(max(int(self.min_table), 0), 150)),
                uniform=group,
            )
            if self.compact is not None:
                available = max(0, total - int(self.bar.winfo_reqheight()))
                self.compact(int(round(available * (1.0 - self.table_fraction))))
            return

        # Release the old fixed-pixel table constraint. Controls keep their
        # natural/requested height; the table receives only the remaining space.
        self.table_height = None
        self.parent.rowconfigure(self.absorb_row, weight=0)
        self.parent.rowconfigure(
            self.table_row,
            weight=1,
            minsize=min(max(int(self.min_table), 0), 90),
        )


def _hide_label_with_text(root, needle: str, hide: bool) -> None:
    for child in root.winfo_children():
        try:
            text = str(child.cget("text"))
        except tk.TclError:
            text = ""
        if needle in text:
            if hide:
                child.grid_remove()
            else:
                child.grid()
        _hide_label_with_text(child, needle, hide)


def _move_tree_container(tree: ttk.Treeview, *, row: int) -> None:
    tree.master.grid_configure(row=row)


class ResizableOZPriceAnalyzerApp(OZPriceAnalyzerApp):
    """Static table layout with protected controls and non-draggable separators."""

    def _build_overview_tab(self) -> None:
        super()._build_overview_tab()
        _move_tree_container(self.overview_tree, row=2)
        self.overview_tab.rowconfigure(1, weight=0)
        self.overview_tab.rowconfigure(2, weight=1)

        self._overview_splitter = _GridSplitter(
            self, self.overview_tab, row=1, table_row=2, absorb_row=0,
            min_upper=210, min_table=190,
            protected_rows=(0,), table_fraction=0.5,
            bar=self.overview_controls,
        )

    def _build_scenario_tab(self) -> None:
        super()._build_scenario_tab()
        _move_tree_container(self.scenario_tree, row=5)
        self.scenario_kpi_frame.grid_configure(row=6)
        self.scenario_tab.rowconfigure(4, weight=0)
        self.scenario_tab.rowconfigure(5, weight=1)

        def compact(upper: int):
            _hide_label_with_text(
                self.scenario_tab,
                "Объем продаж остается текущим",
                upper < 175,
            )

        self._scenario_splitter = _GridSplitter(
            self, self.scenario_tab, row=4, table_row=5, absorb_row=3,
            min_upper=135, min_table=220, compact=compact,
        )

    def _build_settings_tab(self) -> None:
        super()._build_settings_tab()
        _move_tree_container(self.products_tree, row=4)
        self.settings_tab.rowconfigure(1, weight=0)
        self.settings_tab.rowconfigure(2, weight=0)
        self.settings_tab.rowconfigure(3, weight=0)
        self.settings_tab.rowconfigure(4, weight=1)

        def compact(upper: int):
            _hide_label_with_text(
                self.settings_tab,
                "При первом запуске справочник пуст",
                upper < 330,
            )
            _hide_label_with_text(
                self.settings_tab,
                "Архив содержит историю расчетов",
                upper < 300,
            )

        self._settings_splitter = _GridSplitter(
            self, self.settings_tab, row=3, table_row=4, absorb_row=1,
            min_upper=265, min_table=185, compact=compact,
            protected_rows=(0, 1, 2),
        )


def run_app() -> None:
    app = ResizableOZPriceAnalyzerApp()
    app.mainloop()
