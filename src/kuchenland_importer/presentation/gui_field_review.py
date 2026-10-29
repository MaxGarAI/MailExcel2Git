"""Main-thread field decisions exchanged with a waiting worker through queues."""

import tkinter as tk
from dataclasses import dataclass, field
from queue import Queue
from tkinter import messagebox, ttk

from kuchenland_importer.domain.field_review import FieldChoice, FieldDecision


@dataclass(slots=True)
class ReviewRequest:
    choice: FieldChoice
    response: Queue[FieldDecision | None] = field(default_factory=Queue)


def show_review(parent: tk.Tk, request: ReviewRequest) -> None:
    choice = request.choice
    window = tk.Toplevel(parent)
    window.title("Новое поле — решение перед импортом")
    window.geometry("760x430")
    window.transient(parent)
    window.grab_set()
    body = ttk.Frame(window, padding=18)
    body.pack(fill="both", expand=True)
    ttk.Label(body, text=f"Поставщик: {choice.supplier}", wraplength=700).pack(anchor="w")
    ttk.Label(body, text="Листы: " + ", ".join(choice.sheets), wraplength=700).pack(anchor="w")
    ttk.Label(body, text=f"Новое поле: {choice.source}", wraplength=700).pack(anchor="w", pady=10)
    ttk.Label(body, text="Примеры: " + " | ".join(choice.samples),
              wraplength=700).pack(anchor="w")
    action = tk.StringVar(value="Пропустить")
    ttk.Combobox(body, textvariable=action, state="readonly", values=[
        "Пропустить", "Связать с существующим", "Создать справа"
    ]).pack(fill="x", pady=10)
    target = tk.StringVar()
    ttk.Label(body, text="Существующий столбец (для связывания):").pack(anchor="w")
    ttk.Combobox(body, textvariable=target, state="readonly",
                 values=choice.targets).pack(fill="x")
    title = tk.StringVar(value=choice.source.removeprefix("source:"))
    ttk.Label(body, text="Название нового столбца (для создания):").pack(anchor="w", pady=(10, 0))
    ttk.Entry(body, textvariable=title).pack(fill="x")
    ttk.Label(body, text="Решение запомнится для этой книги, поставщика и типа листа.").pack(
        anchor="w", pady=10)

    def finish(decision: FieldDecision | None) -> None:
        request.response.put(decision)
        window.destroy()

    def accept() -> None:
        selected = action.get()
        if selected == "Связать с существующим" and target.get() not in choice.targets:
            messagebox.showerror("Поле", "Выберите целевой столбец.", parent=window)
            return
        if selected == "Создать справа" and not title.get().strip():
            messagebox.showerror("Поле", "Введите название столбца.", parent=window)
            return
        finish(FieldDecision(
            {"Пропустить": "skip", "Связать с существующим": "map", "Создать справа": "create"}[
                selected], title.get().strip() if selected == "Создать справа" else target.get()))

    buttons = ttk.Frame(body)
    buttons.pack(fill="x")
    ttk.Button(buttons, text="Применить и запомнить", command=accept).pack(side="left")
    ttk.Button(buttons, text="Отменить импорт", command=lambda: finish(None)).pack(side="right")
    window.protocol("WM_DELETE_WINDOW", lambda: finish(None))
    parent.wait_window(window)
