"""Isolated Excel process for legacy/binary sources; never attach to user Excel."""

from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from kuchenland_importer.domain.cell_values import CellValue, ExcelErrorValue
from kuchenland_importer.domain.workbook import SourceCell, SourceSheet

ERROR_CODES = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
}


def read_legacy(path: Path) -> tuple[SourceSheet, ...]:
    from kuchenland_importer.infrastructure.excel.reader import cell_value

    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app: Any = None
    book: Any = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        app.AskToUpdateLinks = False
        book = app.Workbooks.Open(
            str(path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            Password="__unsupported_encryption__",
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        if any(sheet.Type != -4167 for sheet in book.Sheets):
            raise ValueError("Листы макросов и иные типы листов не поддерживаются.")
        result = []
        for sheet in book.Worksheets:
            used = sheet.UsedRange
            max_row = used.Row + used.Rows.Count - 1
            max_col = used.Column + used.Columns.Count - 1
            if max_row > 100_000 or max_col > 512:
                raise ValueError("Лист превышает лимит размера.")
            rows = []
            merges = set()
            for row in range(1, max_row + 1):
                cells = []
                for col in range(1, max_col + 1):
                    cell = sheet.Cells(row, col)
                    value = cell.Value
                    if (
                        isinstance(value, int)
                        and -2146826288 <= value <= -2146825289
                        and app.WorksheetFunction.IsError(cell)
                    ):
                        error_code = value - (-2146828288)
                        if error_code not in ERROR_CODES:
                            raise ValueError(f"Неподдерживаемая ошибка Excel: {error_code}.")
                        normalized: CellValue = ExcelErrorValue(ERROR_CODES[error_code])
                    else:
                        normalized = cell_value(value)
                    cells.append(SourceCell(normalized, str(cell.NumberFormat)))
                    if cell.MergeCells:
                        area = cell.MergeArea
                        merges.add(
                            (
                                area.Row,
                                area.Column,
                                area.Row + area.Rows.Count - 1,
                                area.Column + area.Columns.Count - 1,
                            )
                        )
                rows.append(tuple(cells))
            result.append(SourceSheet(str(sheet.Name), tuple(rows), tuple(sorted(merges))))
        return tuple(result)
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                book = None
                app = None
                pythoncom.CoUninitialize()
