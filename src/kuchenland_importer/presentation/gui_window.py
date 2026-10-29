"""Main Tk window. Only this thread touches widgets or opens dialogs."""

import os
import tkinter as tk
from dataclasses import replace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from kuchenland_importer import __version__
from kuchenland_importer.app import Application
from kuchenland_importer.application.job_control import Progress
from kuchenland_importer.infrastructure.settings import Settings
from kuchenland_importer.infrastructure.settings_store import save_settings
from kuchenland_importer.presentation.gui_field_review import ReviewRequest, show_review
from kuchenland_importer.presentation.gui_jobs import GuiResult, import_job
from kuchenland_importer.presentation.gui_report import ReportPanel
from kuchenland_importer.presentation.gui_settings import SettingsDialog
from kuchenland_importer.presentation.gui_worker import Completed, Failed, GuiWorker


class ImportWindow:
    def __init__(self, root: tk.Tk, app: Application, config: Path, personal: Path) -> None:
        self.root, self.app, self.config, self.personal = root, app, config, personal
        self.worker = GuiWorker()
        self.last_result: GuiResult | None = None
        self.closing = False
        root.title(f"Kuchenland — импорт закупок {__version__}")
        root.geometry("1120x760")
        root.minsize(920, 640)
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 19, "bold"))
        body = ttk.Frame(root, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Импорт закупок", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text="Выделите письма в классическом Outlook и закройте общую книгу "
            "в Excel перед импортом.",
        ).pack(anchor="w", pady=(6, 18))
        workbook = ttk.Frame(body)
        workbook.pack(fill="x")
        ttk.Label(workbook, text="Общая книга").pack(side="left", padx=(0, 12))
        self.workbook = tk.StringVar(value=str(app.settings.workbook or ""))
        self.entry = ttk.Entry(workbook, textvariable=self.workbook)
        self.entry.pack(side="left", fill="x", expand=True)
        self.actions: list[ttk.Button] = []
        self._button(workbook, "Выбрать…", self.choose_workbook)
        commands = ttk.Frame(body)
        commands.pack(fill="x", pady=14)
        self._button(commands, "Импорт из Outlook", self.start_outlook)
        self._button(commands, "Из сохранённых вложений…", self.start_manifest)
        self._button(commands, "Сезоны и вкладки…", self.settings)
        self.cancel_button = ttk.Button(
            commands, text="Отменить", command=self.cancel, state="disabled"
        )
        self.cancel_button.pack(side="right")
        self.status = tk.StringVar(value="Готово к работе")
        ttk.Label(body, textvariable=self.status, wraplength=1000).pack(anchor="w", pady=(2, 6))
        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 16))
        self.report = ReportPanel(body)
        self.report.pack(fill="both", expand=True)
        bottom = ttk.Frame(body)
        bottom.pack(fill="x", pady=(16, 0))
        for label, command in (
            ("Отчёт", self.open_report),
            ("Журналы", self.open_logs),
            ("Резервные копии", self.open_backups),
        ):
            ttk.Button(bottom, text=label, command=command).pack(side="left", padx=(0, 8))
        options = ttk.Frame(body)
        options.pack(fill="x", pady=(10, 0))
        self._button(options, "Названия столбцов…", self.columns)
        self._button(options, "Правила новых полей…", self.field_rules)
        self._button(options, "Восстановить книгу…", self.restore)
        self._button(options, "Снять аварийную блокировку…", self.recover_lock)
        root.after(100, self.poll)

    def _button(self, parent: ttk.Frame, title: str, command: object) -> None:
        from collections.abc import Callable
        from typing import cast

        button = ttk.Button(parent, text=title, command=cast(Callable[[], object], command))
        button.pack(side="left", padx=(8, 0))
        self.actions.append(button)

    def choose_workbook(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Выберите общую книгу",
            filetypes=[("Книга Excel", "*.xlsx *.xlsm")],
        )
        if path:
            self.workbook.set(path)
            try:
                self._apply(replace(self.app.settings, workbook=Path(path).resolve()))
            except Exception as error:
                messagebox.showerror("Настройки", str(error), parent=self.root)

    def _apply(self, settings: Settings) -> None:
        save_settings(self.personal, settings)
        self.app.settings = settings

    def settings(self) -> None:
        if not self.worker.busy:
            SettingsDialog(self.root, self.app.settings, self._apply)

    def columns(self) -> None:
        from kuchenland_importer.presentation.gui_columns import ColumnsDialog

        try:
            ColumnsDialog(self.root, self.config)
        except (OSError, ValueError) as error:
            messagebox.showerror("Названия столбцов", str(error), parent=self.root)

    def field_rules(self) -> None:
        from kuchenland_importer.presentation.gui_field_rules import show_rules

        if not self.worker.busy:
            show_rules(self.root, self.app.settings.data_dir / "field-rules.json")

    def _prepare(self) -> bool:
        try:
            text = self.workbook.get().strip()
            path = Path(text).resolve() if text else None
            if path is None or not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm"}:
                raise ValueError("Выберите существующую общую книгу XLSX или XLSM.")
            self._apply(replace(self.app.settings, workbook=path))
            return True
        except Exception as error:
            messagebox.showerror("Импорт", str(error), parent=self.root)
            return False

    def start_outlook(self) -> None:
        self._start(None)

    def start_manifest(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Манифест сохранённых вложений",
            filetypes=[("Манифест JSON", "*.json")],
        )
        if path:
            self._start(Path(path))

    def _start(self, manifest: Path | None) -> None:
        if self.worker.busy or not self._prepare():
            return
        self.last_result = None
        self.report.summary.set("Выполняется новый импорт. Результат появится после завершения.")
        self.report.table.delete(*self.report.table.get_children())
        self.worker.start(lambda control: import_job(self.app, self.config, control, manifest))
        self._busy(True)

    def _busy(self, busy: bool) -> None:
        for button in self.actions:
            button.configure(state="disabled" if busy else "normal")
        self.entry.configure(state="disabled" if busy else "normal")
        self.cancel_button.configure(state="normal" if busy else "disabled")
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def poll(self) -> None:
        for event in self.worker.drain():
            if isinstance(event, Progress):
                self.status.set(event.message)
                self.cancel_button.configure(state="normal" if event.cancellable else "disabled")
            elif isinstance(event, ReviewRequest):
                if self.closing:
                    event.response.put(None)
                else:
                    show_review(self.root, event)
            elif isinstance(event, Completed):
                self._busy(False)
                self.last_result = event.result
                self.report.show_result(event.result)
                has_errors = event.result.outcome.report.error_count or any(
                    i.severity.value == "error" for i in event.result.capture_issues
                )
                self.status.set(
                    event.result.operation_message
                    or (
                        "Завершено с замечаниями — проверьте отчёт"
                        if has_errors
                        else "Импорт завершён"
                    )
                )
            elif isinstance(event, Failed):
                self._busy(False)
                self.status.set(event.message)
                self.report.summary.set(event.message)
                if not event.cancelled and not self.closing:
                    messagebox.showerror("Операция не завершена", event.message, parent=self.root)
        if self.closing and not self.worker.busy:
            self.root.destroy()
        else:
            self.root.after(100, self.poll)

    def cancel(self) -> None:
        self.worker.cancel()
        self.status.set("Запрошена отмена. Ожидаем безопасной точки завершения.")
        self.cancel_button.configure(state="disabled")

    def close(self) -> None:
        if self.worker.busy:
            if messagebox.askyesno(
                "Выполняется импорт",
                "Закрыть окно после завершения операции? "
                "До записи будет запрошена отмена; начатая запись завершится.",
                parent=self.root,
            ):
                self.closing = True
                self.worker.cancel()
        else:
            self.root.destroy()

    def _open(self, path: Path | None) -> None:
        if path is None:
            messagebox.showinfo(
                "Файл", "Для последнего запуска файл отчёта недоступен.", parent=self.root
            )
            return
        try:
            os.startfile(path)
        except OSError as error:
            messagebox.showerror("Не удалось открыть", str(error), parent=self.root)

    def open_report(self) -> None:
        self._open(self.last_result.outcome.report_path if self.last_result else None)

    def open_logs(self) -> None:
        self._open(self.app.paths.logs)

    def open_backups(self) -> None:
        self._open(self.app.paths.backup)

    def restore(self) -> None:
        self._recovery(False)

    def recover_lock(self) -> None:
        self._recovery(True)

    def _recovery(self, lock_only: bool) -> None:
        from kuchenland_importer.presentation.gui_recovery import recovery_job

        if self.worker.busy or not self._prepare():
            return
        journal = filedialog.askopenfilename(
            parent=self.root,
            title="Журнал резервной копии",
            initialdir=self.app.paths.backup,
            filetypes=[("Журнал JSON", "*.json")],
        )
        if not journal:
            return
        title = (
            "Снять аварийную блокировку" if lock_only else "Восстановить книгу из резервной копии"
        )
        if not messagebox.askyesno(
            title, f"{title}?\n{self.app.settings.workbook}", parent=self.root
        ):
            return
        self.worker.start(lambda control: recovery_job(self.app, Path(journal), lock_only, control))
        self._busy(True)
