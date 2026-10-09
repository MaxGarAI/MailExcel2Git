"""Keep service/header formulas verbatim despite Excel's automatic row-deletion adjustments."""

from typing import Any

from kuchenland_importer.domain.master import MasterSnapshot


def capture_service(book: Any, master: MasterSnapshot) -> dict[tuple[str, str, bool], str]:
    result = {}
    for name, layout in master.layouts.items():
        sheet = book.Worksheets(name)
        for row in range(1, layout.first_row):
            for column in range(1, max(layout.columns.values()) + 1):
                cell = sheet.Cells(row, column)
                if not cell.HasFormula:
                    continue
                array = bool(cell.HasArray)
                area = cell.CurrentArray if array else cell
                result[(name, str(area.Address), array)] = str(
                    area.FormulaArray if array else cell.Formula2
                )
    return result


def restore_service(book: Any, cells: dict[tuple[str, str, bool], str]) -> None:
    for (name, address, array), formula in cells.items():
        area = book.Worksheets(name).Range(address)
        if array:
            if area.FormulaArray != formula:
                area.FormulaArray = formula
        elif area.Formula2 != formula:
            area.Formula2 = formula
