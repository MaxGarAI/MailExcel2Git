"""Own one Excel STA session and always close without saving on failure."""

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client
from loguru import logger


@contextmanager
def owned_workbook(path: Path) -> Iterator[Any]:
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app: Any = None
    book: Any = None
    failed = False
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        app.AskToUpdateLinks = False
        seed = app.Workbooks.Add()
        try:
            app.Calculation = -4135
            app.CalculateBeforeSave = False
        finally:
            seed.Close(SaveChanges=False)
        book = app.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=False,
            Password="__unsupported_encryption__",
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        app.Calculation = -4135  # Manual: do not recalculate imported values or unrelated cells.
        app.CalculateBeforeSave = False
        if book.ReadOnly or book.ProtectStructure:
            raise ValueError("Рабочая копия защищена или открылась только для чтения.")
        if any(sheet.Type != -4167 for sheet in book.Sheets):
            raise ValueError("Листы макросов не допускаются в общей книге.")
        yield book
    except BaseException:
        failed = True
        raise
    finally:
        cleanup_errors: list[Exception] = []
        for resource, method, arguments in (
            (book, "Close", {"SaveChanges": False}),
            (app, "Quit", {}),
        ):
            if resource is not None:
                try:
                    getattr(resource, method)(**arguments)
                except Exception as error:
                    cleanup_errors.append(error)
                    logger.warning("Не удалось выполнить Excel.{}: {}", method, error)
        book = app = None
        pythoncom.CoUninitialize()
        if cleanup_errors and not failed:
            raise cleanup_errors[0]
