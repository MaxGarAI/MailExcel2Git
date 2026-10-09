"""Capture absent fields by role, delete old rows bottom-up, and append replacement rows."""

from dataclasses import dataclass
from typing import Any

from kuchenland_importer.domain.master import MasterSnapshot, SheetLayout
from kuchenland_importer.domain.write_plan import WritePlan
from kuchenland_importer.infrastructure.excel.com_values import write_value
from kuchenland_importer.infrastructure.excel.formula_routes import rebind_formula


@dataclass(frozen=True, slots=True)
class PreservedCell:
    value: Any
    formula_r1c1: str | None
    number_format: str


def capture_rows(
    book: Any, master: MasterSnapshot, plan: WritePlan
) -> dict[tuple[str, bool], dict[str, PreservedCell]]:
    preserved = {}
    for action in plan.products:
        if action.previous is None:
            continue
        for old in action.previous.rows:
            layout = master.layouts[old.sheet]
            sheet = book.Worksheets(old.sheet)
            fields = {}
            for key, column in layout.columns.items():
                if key == "photo":
                    continue
                cell = sheet.Cells(old.row, column)
                fields[key] = PreservedCell(
                    cell.Value,
                    str(cell.FormulaR1C1) if cell.HasFormula else None,
                    str(cell.NumberFormatLocal),
                )
            preserved[(action.incoming.product.article, "photo" in layout.columns)] = fields
    return preserved


def delete_rows(book: Any, master: MasterSnapshot, plan: WritePlan) -> None:
    by_sheet: dict[str, set[int]] = {}
    for action in plan.products:
        if action.previous:
            for old in action.previous.rows:
                by_sheet.setdefault(old.sheet, set()).add(old.row)
    for name, rows in by_sheet.items():
        sheet = book.Worksheets(name)
        layout = master.layouts[name]
        # Explicitly remove floating pictures anchored to replaced product rows.
        for index in range(sheet.Shapes.Count, 0, -1):
            shape = sheet.Shapes(index)
            if int(shape.TopLeftCell.Row) in rows:
                shape.Delete()
        for row in sorted(rows, reverse=True):
            sheet.Rows(row).Delete()
        layout.last_row -= len(rows)


def append_row(sheet: Any, layout: SheetLayout) -> int:
    last = int(sheet.Cells(sheet.Rows.Count, layout.columns["article"]).End(-4162).Row)
    row = max(layout.first_row, last + 1)
    if row >= 100_000:
        raise ValueError("Общая книга достигла лимита строк.")
    # Preserve a footer/template row below the last article instead of clearing it.
    area = sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, max(layout.columns.values())))
    if sheet.Application.WorksheetFunction.CountA(area):
        sheet.Rows(row).Insert()
    if last >= layout.first_row:
        # Formats only: copying another product's formulas/values would be incorrect.
        sheet.Rows(last).Copy()
        sheet.Rows(row).PasteSpecial(Paste=-4122)
        sheet.Application.CutCopyMode = False
        sheet.Rows(row).RowHeight = sheet.Rows(last).RowHeight
    sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, len(layout.columns))).ClearContents()
    layout.last_row = max(layout.last_row, row)
    return row


def restore_absent(
    sheet: Any,
    row: int,
    layout: SheetLayout,
    fields: dict[str, PreservedCell],
    incoming: set[str],
    *,
    preserve_manual: bool,
    formula_routes: dict[str, str] | None = None,
) -> None:
    for key, prior in fields.items():
        if key in incoming or key not in layout.columns or key == "photo":
            continue
        cell = sheet.Cells(row, layout.columns[key])
        cell.NumberFormatLocal = prior.number_format
        if prior.formula_r1c1:
            cell.FormulaR1C1 = rebind_formula(prior.formula_r1c1, formula_routes or {})
        elif preserve_manual:
            # Value may be an Excel error; convert HRESULT to the domain error explicitly.
            from kuchenland_importer.domain.cell_values import ExcelErrorValue
            from kuchenland_importer.infrastructure.excel.legacy_reader import ERROR_CODES

            value = prior.value
            if isinstance(value, int) and value + 2146828288 in ERROR_CODES:
                value = ExcelErrorValue(ERROR_CODES[value + 2146828288])
            write_value(cell, value)
