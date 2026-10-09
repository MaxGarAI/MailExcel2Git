"""Convert an old-format source to a disposable OOXML copy without saving the original."""

from pathlib import Path
from typing import Any


def convert_photo_copy(source: Path, destination: Path) -> None:
    import pythoncom
    import win32com.client

    if destination.exists():
        raise ValueError("Рабочая копия для фото уже существует.")
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app: Any = None
    book: Any = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        book = app.Workbooks.Open(
            str(source.resolve()), ReadOnly=True, UpdateLinks=0,
            Password="__unsupported_encryption__", AddToMru=False,
        )
        if any(book.Sheets(i).Type != -4167 for i in range(1, book.Sheets.Count + 1)):
            raise ValueError("Книга с листами макросов/диаграмм не поддерживается.")
        book.SaveAs(str(destination.resolve()), FileFormat=51)
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
