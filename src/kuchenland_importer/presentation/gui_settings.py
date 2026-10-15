"""Modal season editor validates the complete settings object before applying."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from dataclasses import replace
from tkinter import messagebox, ttk

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.domain.season import SeasonRoute
from kuchenland_importer.infrastructure.settings import Settings


class SettingsDialog(tk.Toplevel):
    def __init__(
        self, parent: tk.Tk, settings: Settings, apply: Callable[[Settings], None]
    ) -> None:
        super().__init__(parent)
        self.title("Сезоны и вкладки")
        self.geometry("920x540")
        self.minsize(760, 480)
        self.transient(parent)
        self.grab_set()
        self.settings, self.apply = settings, apply
        self.routes = list(settings.routes)
        self.selected: int | None = None
        self.variables = [tk.StringVar() for _ in range(4)]
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(
            body,
            text="Названия вкладок должны совпадать с общей книгой. "
            "Недостающие вкладки создайте в Excel вручную.",
            wraplength=850,
        ).pack(anchor="w")
        self.table = ttk.Treeview(
            body,
            columns=("id", "aliases", "calc", "photo"),
            show="headings",
            height=8,
            selectmode="browse",
        )
        for key, title in (
            ("id", "Код"),
            ("aliases", "Сезоны"),
            ("calc", "Расчётная вкладка"),
            ("photo", "Вкладка ФОТО"),
        ):
            self.table.heading(key, text=title)
            self.table.column(key, width=200, minwidth=80)
        self.table.pack(fill="both", expand=True, pady=12)
        self.table.bind("<<TreeviewSelect>>", self._select)
        form = ttk.Frame(body)
        form.pack(fill="x")
        labels = ("Код (латиница)", "Названия сезонов через ;", "Расчётная вкладка", "Вкладка ФОТО")
        for row, (label, variable) in enumerate(zip(labels, self.variables, strict=True)):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(form, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=10)
        form.columnconfigure(1, weight=1)
        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(12, 0))
        for label, command in (
            ("Новый", self._new),
            ("Применить строку", self._update),
            ("Удалить строку", self._delete),
            ("Сохранить", self._save),
            ("Закрыть", self.destroy),
        ):
            ttk.Button(buttons, text=label, command=command).pack(side="left", padx=(0, 8))
        self._refresh()

    def _refresh(self) -> None:
        self.table.delete(*self.table.get_children())
        for index, route in enumerate(self.routes):
            self.table.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    route.id,
                    "; ".join(route.aliases),
                    route.calculation_sheet,
                    route.photo_sheet,
                ),
            )

    def _select(self, _event: tk.Event[tk.Misc]) -> None:
        selected = self.table.selection()
        if selected:
            self.selected = int(selected[0])
            route = self.routes[self.selected]
            values = (
                route.id,
                "; ".join(route.aliases),
                route.calculation_sheet,
                route.photo_sheet,
            )
            for variable, value in zip(self.variables, values, strict=True):
                variable.set(value)

    def _new(self) -> None:
        self.selected = None
        self.table.selection_remove(*self.table.selection())
        for variable in self.variables:
            variable.set("")

    def _update(self) -> bool:
        try:
            code, aliases, calc, photo = (v.get().strip() for v in self.variables)
            route = SeasonRoute(
                code, tuple(a.strip() for a in aliases.split(";") if a.strip()), calc, photo
            )
            routes = list(self.routes)
            if self.selected is None:
                routes.append(route)
            else:
                routes[self.selected] = route
            replace(self.settings, routes=tuple(routes))
            self.routes = routes
            self._refresh()
            self._new()
            return True
        except (ValueError, ConfigurationError) as error:
            messagebox.showerror("Настройки", str(error), parent=self)
        return False

    def _delete(self) -> None:
        if self.selected is not None:
            del self.routes[self.selected]
            self._refresh()
            self._new()

    def _save(self) -> None:
        # Save pending form edits too; never silently discard a route being edited.
        if any(v.get().strip() for v in self.variables) and not self._update():
            return
        try:
            self.apply(replace(self.settings, routes=tuple(self.routes)))
        except Exception as error:
            messagebox.showerror("Не удалось сохранить", str(error), parent=self)
            return
        self.destroy()
