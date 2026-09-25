from __future__ import annotations

from dataclasses import dataclass
import tkinter as tk
from tkinter import ttk
from typing import Callable

from .resizable_layout import ResizableOZPriceAnalyzerApp, _GridSplitter, _hide_label_with_text


UI_SCALE_LABELS = {
    "Авто (рекомендуется)": "auto",
    "100%": "1.0",
    "90%": "0.9",
    "80%": "0.8",
}
UI_SCALE_VALUES = {value: label for label, value in UI_SCALE_LABELS.items()}


def resolve_ui_scale(preference: str, screen_height: int) -> float:
    """Resolve saved UI scale to an effective factor.

    In auto mode the thresholds target common laptop work areas, including
    Windows display scaling where Tk may report a reduced logical height.
    """
    normalized = str(preference or "auto").strip().casefold()
    explicit = {
        "1": 1.0,
        "1.0": 1.0,
        "100%": 1.0,
        "0.9": 0.9,
        "90%": 0.9,
        "0.8": 0.8,
        "80%": 0.8,
    }
    if normalized in explicit:
        return explicit[normalized]
    if screen_height <= 800:
        return 0.8
    if screen_height <= 950:
        return 0.9
    return 1.0


def fitted_window_size(screen_width: int, screen_height: int) -> tuple[int, int, int, int]:
    """Return width, height, min_width and min_height that fit the screen."""
    width = min(1540, max(1000, screen_width - 48))
    height = min(920, max(620, screen_height - 80))
    min_width = min(1180, max(900, screen_width - 140))
    min_height = min(720, max(560, screen_height - 160))
    return width, height, min_width, min_height


def _grid_child_at_row(parent: tk.Misc, row: int) -> tk.Misc | None:
    for child in parent.winfo_children():
        if child.winfo_manager() != "grid":
            continue
        try:
            if int(child.grid_info().get("row", -1)) == row:
                return child
        except (TypeError, ValueError, tk.TclError):
            continue
    return None


def _button_with_text(root: tk.Misc, needle: str) -> ttk.Button | None:
    for child in root.winfo_children():
        if isinstance(child, ttk.Button):
            try:
                if needle in str(child.cget("text")):
                    return child
            except tk.TclError:
                pass
        found = _button_with_text(child, needle)
        if found is not None:
            return found
    return None


class _DetachedTableWindow:
    """A maximizable synchronized copy of one of the main application tables."""

    def __init__(
        self,
        owner: "LaptopFriendlyOZPriceAnalyzerApp",
        source: ttk.Treeview,
        title: str,
        *,
        action_label: str | None = None,
        action: Callable[[], None] | None = None,
    ):
        self.owner = owner
        self.source = source
        self.action = action
        self._signature: tuple[object, ...] | None = None
        self._closed = False

        self.window = tk.Toplevel(owner)
        self.window.title(f"{title} — отдельная таблица")
        self.window.geometry("1400x800")
        self.window.minsize(900, 500)
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        self.window.bind("<Escape>", lambda _event: self.close())
        self.window.bind("<F5>", lambda _event: self.refresh(force=True))

        toolbar = ttk.Frame(self.window, padding=(10, 8))
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(0, weight=1)
        ttk.Label(
            toolbar,
            text="Таблица синхронизируется с основным окном автоматически.",
            style="Muted.TLabel",
        ).grid(row=0, column=0, sticky="w")
        if action_label and action:
            ttk.Button(toolbar, text=action_label, command=self._run_action).grid(
                row=0, column=1, padx=(8, 0)
            )
        ttk.Button(toolbar, text="Обновить", command=lambda: self.refresh(force=True)).grid(
            row=0, column=2, padx=(8, 0)
        )
        ttk.Button(toolbar, text="Закрыть", command=self.close).grid(
            row=0, column=3, padx=(8, 0)
        )

        container = ttk.Frame(self.window)
        container.grid(row=1, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(1, weight=1)

        columns = tuple(str(column) for column in source.cget("columns"))
        self.tree = ttk.Treeview(container, columns=columns, show="headings")
        xscroll = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        yscroll = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        for column in columns:
            heading = source.heading(column)
            source_column = source.column(column)
            self.tree.heading(column, text=heading.get("text", ""))
            self.tree.column(
                column,
                width=max(int(source_column.get("width", 120)), 70),
                minwidth=max(int(source_column.get("minwidth", 70)), 50),
                stretch=bool(source_column.get("stretch", False)),
                anchor=source_column.get("anchor", "center"),
            )

        self.tree.bind("<<TreeviewSelect>>", self._sync_selection_to_source)
        if action is not None:
            self.tree.bind("<Double-1>", lambda _event: self._run_action())

        self.refresh(force=True)
        self.window.after_idle(self._maximize)
        self.window.after(700, self._poll)

    def _maximize(self) -> None:
        try:
            self.window.state("zoomed")
        except tk.TclError:
            pass

    def _rows_signature(self) -> tuple[object, ...]:
        rows: list[object] = []
        for iid in self.source.get_children(""):
            item = self.source.item(iid)
            rows.append((iid, tuple(item.get("values", ())), tuple(item.get("tags", ()))))
        return tuple(rows)

    def _source_display_columns(self) -> tuple[str, ...]:
        configured = self.source.cget("displaycolumns")
        if configured in ("#all", ("#all",)):
            return tuple(str(column) for column in self.source.cget("columns"))
        if isinstance(configured, (tuple, list)):
            return tuple(str(column) for column in configured)
        return tuple(str(column) for column in self.source.tk.splitlist(configured))

    def refresh(self, *, force: bool = False) -> None:
        if self._closed or not self.window.winfo_exists():
            return
        display_columns = self._source_display_columns()
        rows = self._rows_signature()
        signature = (display_columns, rows)
        if not force and signature == self._signature:
            return

        selected = tuple(self.tree.selection())
        yview = self.tree.yview()
        self.tree.configure(displaycolumns=display_columns)
        self.tree.delete(*self.tree.get_children(""))
        for iid, values, tags in rows:
            self.tree.insert("", "end", iid=str(iid), values=values, tags=tags)
        self.owner._configure_value_tags(self.tree)

        valid = [iid for iid in selected if self.tree.exists(iid)]
        if valid:
            self.tree.selection_set(valid)
        if yview:
            self.tree.yview_moveto(yview[0])
        self._signature = signature

    def _poll(self) -> None:
        if self._closed or not self.window.winfo_exists():
            return
        self.refresh()
        self.window.after(700, self._poll)

    def _sync_selection_to_source(self, _event=None) -> None:
        selection = [iid for iid in self.tree.selection() if self.source.exists(iid)]
        if not selection:
            return
        self.source.selection_set(selection)
        self.source.focus(selection[0])
        self.source.see(selection[0])

    def _run_action(self) -> None:
        self._sync_selection_to_source()
        if self.action is not None:
            self.action()
            self.owner.after_idle(lambda: self.refresh(force=True))

    def focus(self) -> None:
        if self._closed or not self.window.winfo_exists():
            return
        self.window.deiconify()
        self.window.lift()
        self.window.focus_force()

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            self.window.destroy()
        except tk.TclError:
            pass


@dataclass
class _TableModeController:
    owner: "LaptopFriendlyOZPriceAnalyzerApp"
    key: str
    title: str
    tab: ttk.Frame
    tree: ttk.Treeview
    splitter: _GridSplitter
    action_label: str | None = None
    action: Callable[[], None] | None = None

    def __post_init__(self) -> None:
        self.fullscreen = False
        self._hidden: list[tk.Misc] = []
        self._normal_grid = dict(self.tree.master.grid_info())
        self._detached: _DetachedTableWindow | None = None

        self.fullscreen_button = self.splitter.add_button("⛶ На весь экран", self.toggle_fullscreen)
        self.splitter.add_button("↗ Отдельно", self.open_detached)

        self.full_toolbar = ttk.Frame(self.tab, style="FloatingTools.TFrame", padding=(6, 4))
        ttk.Button(
            self.full_toolbar,
            text="Вернуть обычный вид",
            style="TableTool.TButton",
            command=self.toggle_fullscreen,
        ).grid(row=0, column=0, padx=2)
        ttk.Button(
            self.full_toolbar,
            text="Открыть отдельно",
            style="TableTool.TButton",
            command=self.open_detached,
        ).grid(row=0, column=1, padx=2)

    def toggle_fullscreen(self) -> None:
        if self.fullscreen:
            self.restore()
        else:
            self.expand()

    def expand(self) -> None:
        if self.fullscreen:
            return
        self.owner._restore_other_table_modes(self.key)
        self._normal_grid = dict(self.tree.master.grid_info())
        self._hidden = [
            child
            for child in self.tab.winfo_children()
            if child is not self.tree.master
            and child.winfo_manager() == "grid"
            and (self.key != "overview" or child is not self.splitter.bar)
        ]
        for child in self._hidden:
            child.grid_remove()

        self.splitter.suspended = True
        for row in range(0, max(self.splitter.table_row + 3, 9)):
            self.tab.rowconfigure(row, weight=0, minsize=0)
        if self.key == "overview":
            # Retain filters and controls on their own row above the enlarged table.
            # The regular 50/50 layout pairs rows 0 and 2 with a grid uniform.
            # Clear it while the summary in row 0 is hidden; otherwise Tk can
            # reserve half the height for an invisible row after DPI changes.
            self.tab.rowconfigure(self.splitter.absorb_row, uniform="")
            self.tab.rowconfigure(self.splitter.table_row, uniform="")
            self.tree.master.grid_configure(row=self.splitter.table_row, column=0, sticky="nsew")
            self.tab.rowconfigure(self.splitter.row, minsize=self.splitter.bar.winfo_reqheight())
            self.tab.rowconfigure(self.splitter.table_row, weight=1, minsize=110)
            self.fullscreen_button.configure(text="Вернуть обычный вид")
        else:
            self.tree.master.grid_configure(row=0, column=0, sticky="nsew")
            self.tab.rowconfigure(0, weight=1)
            self.full_toolbar.place(relx=1.0, x=-10, y=8, anchor="ne")
            self.full_toolbar.lift()
        self.fullscreen = True
        self.owner.status_var.set(
            f"{self.title}: таблица развернута. F11 или Esc — вернуть обычный вид."
        )

    def restore(self) -> None:
        if not self.fullscreen:
            return
        if self.key == "overview":
            self.fullscreen_button.configure(text="⛶ На весь экран")
        else:
            self.full_toolbar.place_forget()
        self.tree.master.grid_configure(
            row=int(self._normal_grid.get("row", self.splitter.table_row)),
            column=int(self._normal_grid.get("column", 0)),
            rowspan=int(self._normal_grid.get("rowspan", 1)),
            columnspan=int(self._normal_grid.get("columnspan", 1)),
            sticky=self._normal_grid.get("sticky", "nsew"),
            padx=self._normal_grid.get("padx", 0),
            pady=self._normal_grid.get("pady", 0),
            ipadx=self._normal_grid.get("ipadx", 0),
            ipady=self._normal_grid.get("ipady", 0),
        )
        for child in self._hidden:
            try:
                child.grid()
            except tk.TclError:
                pass
        self._hidden.clear()
        for row in range(0, max(self.splitter.table_row + 3, 9)):
            self.tab.rowconfigure(row, weight=0, minsize=0)
        self.splitter.suspended = False
        self.splitter._apply()
        self.fullscreen = False
        self.owner.status_var.set(self.owner._current_run_status())

    def open_detached(self) -> None:
        if self._detached is not None and not self._detached._closed:
            self._detached.focus()
            return
        self._detached = _DetachedTableWindow(
            self.owner,
            self.tree,
            self.title,
            action_label=self.action_label,
            action=self.action,
        )


class LaptopFriendlyOZPriceAnalyzerApp(ResizableOZPriceAnalyzerApp):
    """v0.5.4 display modes for laptops and smaller screens."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._base_tk_scaling = float(self.tk.call("tk", "scaling"))
        self._splitter_base_upper = {
            "overview": self._overview_splitter.min_upper,
            "scenario": self._scenario_splitter.min_upper,
            "settings": self._settings_splitter.min_upper,
        }
        self._table_modes: dict[str, _TableModeController] = {}

        self._configure_display_styles()
        self._install_scale_control()
        self._install_table_modes()
        self._fit_to_screen()
        self._apply_saved_ui_scale()
        self.bind("<F11>", self._toggle_current_table_mode)
        self.bind("<Escape>", self._escape_table_mode)

    def _configure_display_styles(self) -> None:
        style = ttk.Style(self)
        style.configure("TableTool.TButton", padding=(8, 3))
        style.configure("FloatingTools.TFrame", relief="solid", borderwidth=1)

    def _install_scale_control(self) -> None:
        settings_frame = _grid_child_at_row(self.settings_tab, 1)
        if settings_frame is None:
            return

        save_button = _button_with_text(settings_frame, "Сохранить настройки")
        if save_button is not None:
            try:
                save_button.grid_configure(columnspan=2)
            except tk.TclError:
                pass

        saved = self.db.get_setting("ui_scale", "auto")
        self.ui_scale_var = tk.StringVar(
            value=UI_SCALE_VALUES.get(saved, "Авто (рекомендуется)")
        )
        ttk.Label(settings_frame, text="Масштаб интерфейса:").grid(
            row=3, column=2, sticky="e", padx=(24, 8), pady=(10, 0)
        )
        combo = ttk.Combobox(
            settings_frame,
            textvariable=self.ui_scale_var,
            state="readonly",
            values=tuple(UI_SCALE_LABELS),
            width=22,
        )
        combo.grid(row=3, column=3, sticky="w", pady=(10, 0))
        combo.bind("<<ComboboxSelected>>", self._on_ui_scale_selected)

    def _install_table_modes(self) -> None:
        self._table_modes = {
            "overview": _TableModeController(
                self,
                "overview",
                "Обзор",
                self.overview_tab,
                self.overview_tree,
                self._overview_splitter,
            ),
            "scenario": _TableModeController(
                self,
                "scenario",
                "Сценарий цены",
                self.scenario_tab,
                self.scenario_tree,
                self._scenario_splitter,
            ),
            "settings": _TableModeController(
                self,
                "settings",
                "Справочник себестоимости",
                self.settings_tab,
                self.products_tree,
                self._settings_splitter,
                action_label="Изменить выбранный",
                action=self.edit_product,
            ),
        }

    def _fit_to_screen(self) -> None:
        screen_width = int(self.winfo_screenwidth())
        screen_height = int(self.winfo_screenheight())
        width, height, min_width, min_height = fitted_window_size(screen_width, screen_height)
        self.minsize(min_width, min_height)
        x = max((screen_width - width) // 2, 0)
        y = max((screen_height - height) // 2, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")
        if screen_height <= 900 or screen_width <= 1440:
            self.after_idle(self._maximize_on_small_screen)

    def _maximize_on_small_screen(self) -> None:
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

    def _saved_scale_preference(self) -> str:
        if hasattr(self, "ui_scale_var"):
            return UI_SCALE_LABELS.get(self.ui_scale_var.get(), "auto")
        return self.db.get_setting("ui_scale", "auto")

    def _apply_saved_ui_scale(self) -> None:
        preference = self._saved_scale_preference()
        factor = resolve_ui_scale(preference, int(self.winfo_screenheight()))
        self._apply_ui_scale(factor)

    def _on_ui_scale_selected(self, _event=None) -> None:
        preference = self._saved_scale_preference()
        self.db.set_setting("ui_scale", preference)
        factor = resolve_ui_scale(preference, int(self.winfo_screenheight()))
        self._apply_ui_scale(factor)
        self.status_var.set(
            f"Масштаб интерфейса: {int(round(factor * 100))}%"
            + (" (авто)" if preference == "auto" else "")
        )

    def _apply_ui_scale(self, factor: float) -> None:
        factor = min(1.0, max(0.8, float(factor)))
        self.tk.call("tk", "scaling", self._base_tk_scaling * factor)
        super()._preview_theme()
        self._apply_density_styles(factor)

        for key, splitter in (
            ("overview", self._overview_splitter),
            ("scenario", self._scenario_splitter),
            ("settings", self._settings_splitter),
        ):
            splitter.min_upper = max(90, int(round(self._splitter_base_upper[key] * factor)))
            splitter._apply()

        compact = factor < 0.99
        _hide_label_with_text(
            self,
            "Отчеты Ozon, история, контроль начислений и плановая доходность",
            compact,
        )
        header = _grid_child_at_row(self, 0)
        if isinstance(header, ttk.Frame):
            try:
                header.configure(
                    padding=(18, 8, 18, 8) if compact else (22, 18, 22, 16)
                )
            except tk.TclError:
                pass
        try:
            self.notebook.grid_configure(
                padx=12 if compact else 18,
                pady=(0, 6 if compact else 12),
            )
        except tk.TclError:
            pass

    def _apply_density_styles(self, factor: float) -> None:
        style = ttk.Style(self)
        scale = lambda value, floor: max(floor, int(round(value * factor)))
        style.configure("Treeview", rowheight=scale(30, 22))
        style.configure(
            "Treeview.Heading",
            padding=(scale(8, 5), scale(8, 4)),
        )
        style.configure(
            "TButton",
            padding=(scale(14, 8), scale(8, 4)),
        )
        style.configure(
            "Accent.TButton",
            padding=(scale(16, 9), scale(9, 5)),
        )
        style.configure("TEntry", padding=scale(7, 4))
        style.configure("TCombobox", padding=scale(6, 4))
        style.configure("TSpinbox", padding=scale(6, 4))
        style.configure(
            "TNotebook.Tab",
            padding=(scale(16, 10), scale(10, 5)),
        )
        style.configure("TableTool.TButton", padding=(scale(8, 5), scale(3, 2)))

    def _preview_theme(self, _event=None) -> None:
        super()._preview_theme(_event)
        if hasattr(self, "_base_tk_scaling"):
            factor = resolve_ui_scale(
                self._saved_scale_preference(),
                int(self.winfo_screenheight()),
            )
            self._apply_density_styles(factor)

    def save_settings(self) -> None:
        super().save_settings()
        if hasattr(self, "_base_tk_scaling"):
            self._apply_saved_ui_scale()

    def _reload_settings_after_restore(self) -> None:
        super()._reload_settings_after_restore()
        if hasattr(self, "ui_scale_var"):
            saved = self.db.get_setting("ui_scale", "auto")
            self.ui_scale_var.set(UI_SCALE_VALUES.get(saved, "Авто (рекомендуется)"))
        if hasattr(self, "_base_tk_scaling"):
            self._apply_saved_ui_scale()

    def _restore_other_table_modes(self, current_key: str) -> None:
        for key, controller in self._table_modes.items():
            if key != current_key and controller.fullscreen:
                controller.restore()

    def _current_table_controller(self) -> _TableModeController | None:
        selected = self.notebook.select()
        for controller in self._table_modes.values():
            if str(controller.tab) == selected:
                return controller
        return None

    def _toggle_current_table_mode(self, _event=None):
        controller = self._current_table_controller()
        if controller is not None:
            controller.toggle_fullscreen()
            return "break"
        return None

    def _escape_table_mode(self, _event=None):
        for controller in self._table_modes.values():
            if controller.fullscreen:
                controller.restore()
                return "break"
        return None


def run_app() -> None:
    app = LaptopFriendlyOZPriceAnalyzerApp()
    app.mainloop()
