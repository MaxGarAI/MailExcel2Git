"""Edit aliases without changing canonical product fields."""

import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from kuchenland_importer.infrastructure.column_settings import column_path, read_fields, save_fields


class ColumnsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, config: Path) -> None:
        super().__init__(parent)
        self.title("Названия столбцов поставщиков")
        self.geometry("850x620")
        self.transient(parent)
        self.grab_set()
        self.destination = config.parent / "columns.local.toml"
        self.fields = read_fields(column_path(config))
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(
            body,
            text="Добавляйте названия через ;. Одно название может относиться "
            "только к одному полю.",
            wraplength=780,
        ).pack(anchor="w", pady=(0, 12))
        self.table = ttk.Treeview(body, columns=("field", "names"), show="headings")
        self.table.heading("field", text="Поле товара")
        self.table.heading("names", text="Названия во вложениях")
        self.table.column("field", width=150, stretch=False)
        self.table.column("names", width=600)
        scroll = ttk.Scrollbar(body, command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.table.pack(fill="both", expand=True)
        self.name = tk.StringVar()
        self.names = tk.StringVar()
        ttk.Label(body, textvariable=self.name).pack(anchor="w", pady=(12, 3))
        ttk.Entry(body, textvariable=self.names).pack(fill="x")
        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=12)
        ttk.Button(buttons, text="Применить строку", command=self._update).pack(side="left")
        ttk.Button(buttons, text="Сохранить", command=self._save).pack(side="right")
        self.table.bind("<<TreeviewSelect>>", lambda event: self._select())
        self._refresh()

    def _refresh(self) -> None:
        self.table.delete(*self.table.get_children())
        for key, values in self.fields.items():
            self.table.insert("", "end", iid=key, values=(key, "; ".join(values)))

    def _select(self) -> None:
        selected = self.table.selection()
        if selected:
            self.name.set(selected[0])
            self.names.set("; ".join(self.fields[selected[0]]))

    def _update(self) -> None:
        key = self.name.get()
        if key:
            self.fields[key] = [
                name.strip() for name in self.names.get().split(";") if name.strip()
            ]
            self._refresh()

    def _save(self) -> None:
        self._update()
        try:
            save_fields(self.destination, self.fields)
        except (OSError, ValueError) as error:
            messagebox.showerror("Названия столбцов", str(error), parent=self)
            return
        self.destroy()
