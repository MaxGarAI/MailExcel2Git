"""Read-only result and issue table rendered exclusively on the Tk thread."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from kuchenland_importer.presentation.gui_jobs import GuiResult


class ReportPanel(ttk.Frame):
    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.summary = tk.StringVar(value="Импорт ещё не выполнялся.")
        ttk.Label(self, textvariable=self.summary, wraplength=850).pack(anchor="w", pady=(0, 12))
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True)
        self.table = ttk.Treeview(
            table_frame, columns=("level", "source", "message"), show="headings", height=15
        )
        for name, title, width in (
            ("level", "Уровень", 100),
            ("source", "Источник", 180),
            ("message", "Описание", 600),
        ):
            self.table.heading(name, text=title)
            self.table.column(name, width=width, minwidth=80)
        vertical = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.table.bind("<Double-1>", self._details)

    def show_result(self, result: GuiResult) -> None:
        report = result.outcome.report
        issues = result.capture_issues + report.issues
        errors = sum(i.severity.value == "error" for i in issues)
        warnings = len(issues) - errors
        saved = "Книга сохранена." if result.outcome.committed else "Книга не изменялась."
        self.summary.set(
            f"{saved} Писем с обработанными вложениями: {report.mails_processed}. "
            f"Добавлено: {report.products_added}; обновлено: {report.products_updated}; "
            f"без изменений: {report.products_unchanged}; пропущено: {report.products_skipped}.\n"
            f"Фото добавлено: {report.photos_added}; заменено: {report.photos_replaced}. "
            f"Ошибок: {errors}; предупреждений: {warnings}."
        )
        if result.operation_message:
            self.summary.set(result.operation_message)
        if result.outcome.backup:
            self.summary.set(self.summary.get() + f"\nРезервная копия: {result.outcome.backup}")
        self.table.delete(*self.table.get_children())
        for issue in issues:
            level = "Ошибка" if issue.severity.value == "error" else "Предупреждение"
            self.table.insert(
                "", "end", values=(level, issue.source, f"[{issue.code}] {issue.message}")
            )

    def _details(self, _event: tk.Event[tk.Misc]) -> None:
        selection = self.table.selection()
        if not selection:
            return
        values = self.table.item(selection[0], "values")
        window = tk.Toplevel(self)
        window.title("Подробности сообщения")
        window.geometry("740x300")
        text = tk.Text(window, wrap="word", padx=12, pady=12)
        text.pack(fill="both", expand=True)
        text.insert("1.0", "\n\n".join(str(value) for value in values))
        text.configure(state="disabled")
