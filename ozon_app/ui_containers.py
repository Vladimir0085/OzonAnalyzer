from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def wrapped_positions(
    sizes: list[tuple[int, int]], available_width: int, gap: int = 8, row_gap: int = 4
) -> tuple[list[tuple[int, int]], int]:
    """Wrap whole control groups without splitting labels from their inputs."""
    width = max(1, available_width)
    positions = []
    x = y = row_height = 0
    for item_width, item_height in sizes:
        if x and x + item_width > width:
            x = 0
            y += row_height + row_gap
            row_height = 0
        positions.append((x, y))
        x += item_width + gap
        row_height = max(row_height, item_height)
    return positions, y + row_height


class WrappingToolbar(ttk.Frame):
    """A naturally sized toolbar that reflows at the actual window/DPI width."""

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent, height=1)
        self.groups: list[ttk.Frame] = []
        self._layout_pending = False
        self.bind("<Configure>", self._schedule_layout)

    def add_group(self) -> ttk.Frame:
        group = ttk.Frame(self)
        self.groups.append(group)
        group.bind("<Configure>", self._schedule_layout)
        group.place(x=0, y=0)
        self._schedule_layout()
        return group

    def _schedule_layout(self, _event=None) -> None:
        if not self._layout_pending:
            self._layout_pending = True
            self.after_idle(self._arrange)

    def _arrange(self) -> None:
        self._layout_pending = False
        if not self.winfo_exists():
            return
        sizes = [(group.winfo_reqwidth(), group.winfo_reqheight()) for group in self.groups]
        positions, height = wrapped_positions(sizes, self.winfo_width())
        for group, (x, y) in zip(self.groups, positions):
            group.place_configure(x=x, y=y)
        height = max(1, height)
        if self.winfo_reqheight() != height:
            self.configure(height=height)


def summary_scrollbars(
    width: int, height: int, content_width: int, content_height: int,
    scrollbar_width: int, scrollbar_height: int,
) -> tuple[bool, bool]:
    """Account for one scrollbar reducing space available on the other axis."""
    horizontal = vertical = False
    for _ in range(3):
        horizontal = content_width > max(1, width - (scrollbar_width if vertical else 0))
        vertical = content_height > max(1, height - (scrollbar_height if horizontal else 0))
    return horizontal, vertical


class ScrollableSummary(ttk.Frame):
    """Keep summary content reachable without taking height from the table."""

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(self, width=1, height=1, highlightthickness=0, borderwidth=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.content = ttk.Frame(self.canvas)
        self.content.columnconfigure(0, weight=1)
        self._window = self.canvas.create_window(0, 0, window=self.content, anchor="nw")
        self.xscroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.yscroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.xscroll.set, yscrollcommand=self.yscroll.set)
        self._wheel_widgets: set[tk.Misc] = set()
        self._sync_pending = False
        self.bind("<Configure>", self._schedule_sync)
        self.canvas.bind("<Configure>", self._schedule_sync)
        self.content.bind("<Configure>", self._schedule_sync)
        self.bind("<<ThemeChanged>>", self._schedule_sync)
        self._schedule_sync()

    def _schedule_sync(self, _event=None) -> None:
        if not self._sync_pending:
            self._sync_pending = True
            self.after_idle(self._sync)

    def _sync(self) -> None:
        self._sync_pending = False
        if not self.winfo_exists():
            return
        self.canvas.configure(background=ttk.Style(self).lookup("TFrame", "background"))
        width, height = self.content.winfo_reqwidth(), self.content.winfo_reqheight()
        horizontal, vertical = summary_scrollbars(
            self.winfo_width(), self.winfo_height(), width, height,
            self.yscroll.winfo_reqwidth(), self.xscroll.winfo_reqheight(),
        )
        if horizontal:
            self.xscroll.grid(row=1, column=0, sticky="ew")
        else:
            self.xscroll.grid_remove()
            self.canvas.xview_moveto(0)
        if vertical:
            self.yscroll.grid(row=0, column=1, sticky="ns")
        else:
            self.yscroll.grid_remove()
            self.canvas.yview_moveto(0)
        viewport_width = self.winfo_width() - (self.yscroll.winfo_reqwidth() if vertical else 0)
        width = max(width, viewport_width)
        self.canvas.itemconfigure(self._window, width=width)
        self.canvas.configure(scrollregion=(0, 0, width, height))
        self._bind_wheel(self.content)

    def _bind_wheel(self, widget: tk.Misc) -> None:
        if widget not in self._wheel_widgets:
            widget.bind("<MouseWheel>", self._scroll, add="+")
            widget.bind("<Button-4>", self._scroll, add="+")
            widget.bind("<Button-5>", self._scroll, add="+")
            self._wheel_widgets.add(widget)
        for child in widget.winfo_children():
            self._bind_wheel(child)

    def _scroll(self, event) -> str | None:
        if self.canvas.yview() == (0.0, 1.0):
            return None
        delta = getattr(event, "delta", 0)
        if not delta and getattr(event, "num", 0) not in (4, 5):
            return None
        direction = -1 if delta > 0 or getattr(event, "num", 0) == 4 else 1
        self.canvas.yview_scroll(direction * max(1, abs(int(delta / 120))), "units")
        return "break"
