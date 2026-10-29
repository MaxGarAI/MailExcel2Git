"""Inspect and forget persisted field decisions without exposing sample product data."""

import json
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from kuchenland_importer.infrastructure.field_rules import load_rules, save_rules


def show_rules(parent: tk.Tk, path: Path) -> None:
    try:
        rules = load_rules(path)
    except (OSError, ValueError, TypeError) as error:
        messagebox.showerror("Правила", str(error), parent=parent)
        return
    window = tk.Toplevel(parent)
    window.title("Сохранённые решения — дополнительные поля")
    window.geometry("1000x450")
    window.transient(parent)
    window.grab_set()
    table = ttk.Treeview(window, columns=("scope", "source", "decision"), show="headings")
    for name, title in (("scope", "Книга / поставщик / тип листа"),
                        ("source", "Поле"), ("decision", "Решение")):
        table.heading(name, text=title)
    table.pack(fill="both", expand=True, padx=12, pady=12)
    keys = list(rules)
    for index, key in enumerate(keys):
        scope = json.loads(key)
        decision = rules[key]
        table.insert("", "end", iid=str(index), values=(
            f"{scope[0]} | {scope[1]} | {'ФОТО' if scope[2] else 'Сезонный'}",
            scope[3], f"{decision.action}: {decision.target}"))

    def forget() -> None:
        selected = table.selection()
        if not selected:
            return
        updated = {k: v for k, v in rules.items()
                   if k not in {keys[int(i)] for i in selected}}
        try:
            save_rules(path, updated)
        except OSError as error:
            messagebox.showerror("Правила", str(error), parent=window)
            return
        rules.clear()
        rules.update(updated)
        for item in selected:
            table.delete(item)

    ttk.Label(window, text="Удалите решение, чтобы согласовать поле при следующем импорте. "
              "Столбцы и данные книги не удаляются.").pack(padx=12)
    ttk.Button(window, text="Забыть выбранные решения", command=forget).pack(pady=12)
    parent.wait_window(window)
