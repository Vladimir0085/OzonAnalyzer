from __future__ import annotations

import tkinter as tk
from typing import Callable


TooltipText = str | Callable[[], str]


class HoverTooltip:
    """Delayed tooltip for a widget and every widget inside it.

    The text may be a callable so that it always describes the current value
    (for example, the report selected in a combobox).
    """

    DELAY_MS = 400
    WRAP_LENGTH = 460

    def __init__(self, widget: tk.Misc, text: TooltipText, *, delay_ms: int | None = None) -> None:
        self.widget = widget
        self._text = text
        self.delay_ms = self.DELAY_MS if delay_ms is None else int(delay_ms)
        self.tipwindow: tk.Toplevel | None = None
        self._after: str | None = None
        self._bound: set[str] = set()
        self.bind_tree(widget)

    def text(self) -> str:
        value = self._text() if callable(self._text) else self._text
        return str(value or "").strip()

    def set_text(self, text: TooltipText) -> None:
        self._text = text

    def bind_tree(self, widget: tk.Misc) -> None:
        """Bind the widget and its current descendants (call again for new children)."""
        key = str(widget)
        if key not in self._bound:
            widget.bind("<Enter>", self._on_enter, add="+")
            widget.bind("<Leave>", self._on_leave, add="+")
            widget.bind("<ButtonPress>", self._on_press, add="+")
            self._bound.add(key)
        for child in widget.winfo_children():
            self.bind_tree(child)

    def _on_enter(self, event=None) -> None:
        if self.tipwindow is not None or self._after is not None:
            return
        x = int(getattr(event, "x_root", 0) or self.widget.winfo_rootx()) + 14
        y = int(getattr(event, "y_root", 0) or self.widget.winfo_rooty()) + 18
        try:
            self._after = self.widget.after(self.delay_ms, lambda: self.show(x, y))
        except tk.TclError:
            self._after = None

    def _on_leave(self, event=None) -> None:
        if event is not None and self._pointer_inside(event):
            return
        self.hide()

    def _on_press(self, _event=None) -> None:
        self.hide()

    def _pointer_inside(self, event) -> bool:
        try:
            target = self.widget.winfo_containing(event.x_root, event.y_root)
        except (tk.TclError, KeyError, AttributeError, TypeError):
            return False
        while target is not None:
            if target is self.widget or str(target) == str(self.widget):
                return True
            target = getattr(target, "master", None)
        return False

    def show(self, x: int | None = None, y: int | None = None) -> None:
        self._after = None
        text = self.text()
        if not text or self.tipwindow is not None:
            return
        try:
            if x is None or y is None:
                x = self.widget.winfo_rootx() + 14
                y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            tip = tk.Toplevel(self.widget)
            tip.withdraw()
            tip.overrideredirect(True)
            try:
                tip.attributes("-topmost", True)
            except tk.TclError:
                pass
            label = tk.Label(
                tip,
                text=text,
                justify="left",
                relief="solid",
                borderwidth=1,
                padx=8,
                pady=5,
                wraplength=self.WRAP_LENGTH,
                background="#fffde7",
                foreground="#111111",
                font="TkDefaultFont",
            )
            label.pack()
            tip.update_idletasks()
            # Keep the tooltip on screen near the right and bottom edges.
            screen_width = tip.winfo_screenwidth()
            screen_height = tip.winfo_screenheight()
            x = max(0, min(int(x), screen_width - tip.winfo_reqwidth() - 4))
            if int(y) + tip.winfo_reqheight() > screen_height:
                y = max(0, int(y) - tip.winfo_reqheight() - 30)
            tip.geometry(f"+{x}+{int(y)}")
            tip.deiconify()
            self.tipwindow = tip
        except tk.TclError:
            self.tipwindow = None

    def hide(self) -> None:
        if self._after is not None:
            try:
                self.widget.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None
        if self.tipwindow is not None:
            try:
                self.tipwindow.destroy()
            except tk.TclError:
                pass
            self.tipwindow = None
