"""Copy bounded cell formats without putting native in-cell pictures on the clipboard."""

from typing import Any

from kuchenland_importer.domain.master import SheetLayout


def copy_row_format(sheet: Any, source_row: int, target_row: int, layout: SheetLayout) -> None:
    last_column = max(layout.columns.values())
    photo_column = layout.columns.get("photo")
    spans = (
        [(1, last_column)]
        if photo_column is None
        else [(1, photo_column - 1), (photo_column + 1, last_column)]
    )
    for first, last in spans:
        if first > last:
            continue
        source = sheet.Range(sheet.Cells(source_row, first), sheet.Cells(source_row, last))
        target = sheet.Range(sheet.Cells(target_row, first), sheet.Cells(target_row, last))
        source.Copy()
        try:
            target.PasteSpecial(Paste=-4122)
        finally:
            sheet.Application.CutCopyMode = False
    if photo_column is not None:
        source = sheet.Cells(source_row, photo_column)
        target = sheet.Cells(target_row, photo_column)
        style = source.Style
        target.Style = style if isinstance(style, str) else style.Name
        for name in ("NumberFormatLocal", "HorizontalAlignment", "VerticalAlignment", "WrapText"):
            setattr(target, name, getattr(source, name))
        for name in ("Color", "PatternColor", "Pattern"):
            setattr(target.Interior, name, getattr(source.Interior, name))
        for border in (7, 8, 9, 10):
            old, new = source.Borders(border), target.Borders(border)
            for name in ("Weight", "Color", "LineStyle"):
                setattr(new, name, getattr(old, name))
    sheet.Rows(target_row).RowHeight = sheet.Rows(source_row).RowHeight
