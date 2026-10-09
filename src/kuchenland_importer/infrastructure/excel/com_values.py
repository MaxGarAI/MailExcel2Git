"""Write real Excel errors and literal supplier text, never evaluate incoming strings."""

from datetime import date, datetime
from decimal import Decimal
from typing import Any

from kuchenland_importer.domain.cell_values import CellValue, ExcelErrorValue

ERROR_FORMULAS = {
    "#NULL!": "=SUM($A$1:$A$1 $B$2:$B$2)",
    "#DIV/0!": "=1/0",
    "#VALUE!": '=VALUE("kuchenland_error")',
    "#REF!": '=INDIRECT("#REF!")',
    "#NAME?": "=kuchenland_unknown_error()",
    "#NUM!": "=SQRT(-1)",
    "#N/A": "=NA()",
}


def international(app: Any, index: int) -> str:
    # Indexed COM properties differ between dynamic Dispatch and generated wrappers.
    dispatch_id = app._oleobj_.GetIDsOfNames("International")
    return str(app._oleobj_.Invoke(dispatch_id, 0, 2, True, index))


def write_value(cell: Any, value: CellValue) -> None:
    if isinstance(value, ExcelErrorValue):
        if value.code not in ERROR_FORMULAS:
            raise ValueError(f"Excel не поддерживает запись ошибки: {value.code}")
        # Excel's automation setter can interpret VT_ERROR as an omitted argument.
        # Calculate an internal constant error, then paste its typed value over the formula.
        cell.Formula = ERROR_FORMULAS[value.code]
        cell.Calculate()
        cell.Copy()
        cell.PasteSpecial(Paste=-4163)
        cell.Application.CutCopyMode = False
        if cell.HasFormula or not cell.Application.WorksheetFunction.IsError(cell):
            raise ValueError("Excel не сохранил ошибку как значение.")
    elif isinstance(value, str):
        previous_format = cell.NumberFormatLocal
        cell.NumberFormatLocal = "@"
        cell.Value2 = value
        cell.NumberFormatLocal = previous_format
        if cell.HasFormula:
            raise ValueError("Входной текст неожиданно стал формулой.")
    elif isinstance(value, Decimal):
        cell.Value2 = float(value)
    elif isinstance(value, date):
        moment = (
            value if isinstance(value, datetime) else datetime.combine(value, datetime.min.time())
        )
        # COM datetime coercion can shift naive dates by the Windows timezone offset.
        # Write the Excel serial directly, respecting the workbook's date system.
        epoch = datetime(1904, 1, 1) if cell.Parent.Parent.Date1904 else datetime(1899, 12, 30)
        cell.Value2 = (moment.replace(tzinfo=None) - epoch).total_seconds() / 86400
        app = cell.Application
        day, month, year = (international(app, code) for code in (21, 20, 19))
        fmt = f"{day * 2}.{month * 2}.{year * 4}"
        if moment.time() != datetime.min.time():
            hour, minute, second = (international(app, code) for code in (22, 23, 24))
            fmt += f" {hour * 2}:{minute * 2}:{second * 2}"
        cell.NumberFormatLocal = fmt
    elif value is None:
        cell.ClearContents()
    else:
        cell.Value2 = value
