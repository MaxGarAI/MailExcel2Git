"""Apply a validated plan solely to an owned working copy through Microsoft Excel."""

from dataclasses import dataclass
from pathlib import Path

from kuchenland_importer.application.target_fields import product_fields
from kuchenland_importer.domain.master import MasterSnapshot
from kuchenland_importer.domain.photos import PhotoAction, photo_action
from kuchenland_importer.domain.write_plan import WritePlan
from kuchenland_importer.infrastructure.excel.calculation_settings import restore_calculation_mode
from kuchenland_importer.infrastructure.excel.com_values import write_value
from kuchenland_importer.infrastructure.excel.incell_writer import InCellPictureWriter
from kuchenland_importer.infrastructure.excel.master_filters import capture_filters, restore_filters
from kuchenland_importer.infrastructure.excel.master_rows import (
    append_row,
    capture_rows,
    delete_rows,
    restore_absent,
)
from kuchenland_importer.infrastructure.excel.master_state import save_state
from kuchenland_importer.infrastructure.excel.owned_session import owned_workbook
from kuchenland_importer.infrastructure.excel.service_cells import capture_service, restore_service


@dataclass(frozen=True, slots=True)
class WrittenProduct:
    article: str
    calculation_sheet: str
    calculation_row: int
    photo_sheet: str
    photo_row: int
    photo_digest: str | None
    photo_action: PhotoAction


def write_master(
    path: Path,
    master: MasterSnapshot,
    plan: WritePlan,
    *,
    preserve_manual: bool,
) -> tuple[WrittenProduct, ...]:
    written = []
    with owned_workbook(path) as book:
        service = capture_service(book, master)
        filters = capture_filters(book, master)
        preserved = capture_rows(book, master, plan)
        delete_rows(book, master, plan)
        for action in plan.products:
            item = action.incoming
            assert item.assignment is not None
            route = item.assignment.route
            rows = []
            old_photo = action.previous.photo if action.previous else None
            decision = photo_action(old_photo, item.product.photo)
            asset = old_photo if decision == PhotoAction.KEEP else item.product.photo
            for name, is_photo in ((route.calculation_sheet, False), (route.photo_sheet, True)):
                layout = master.layouts[name]
                sheet = book.Worksheets(name)
                if sheet.ProtectContents:
                    raise ValueError(f"Лист {name} защищён.")
                fields = product_fields(item, is_photo)
                fields = {key: value for key, value in fields.items() if key in layout.columns}
                row = append_row(sheet, layout)
                restore_absent(
                    sheet,
                    row,
                    layout,
                    preserved.get((item.product.article, is_photo), {}),
                    set(fields),
                    preserve_manual=preserve_manual,
                    formula_routes={
                        (
                            master.layouts[old.sheet].calculation_sheet or old.sheet
                        ): route.calculation_sheet
                        for old in action.previous.rows
                        if "photo" not in master.layouts[old.sheet].columns
                        or master.layouts[old.sheet].calculation_sheet is not None
                    }
                    if action.previous
                    else {},
                )
                for key, value in fields.items():
                    write_value(sheet.Cells(row, layout.columns[key]), value)
                if is_photo and asset is not None:
                    cell = sheet.Cells(row, layout.columns["photo"])
                    if cell.RowHeight < 60:
                        cell.RowHeight = 100
                    InCellPictureWriter().place(sheet, str(cell.Address), asset)
                rows.append(row)
            written.append(
                WrittenProduct(
                    item.product.article,
                    route.calculation_sheet,
                    rows[0],
                    route.photo_sheet,
                    rows[1],
                    asset.pixel_sha256 if asset else None,
                    decision,
                )
            )
        restore_service(book, service)
        restore_filters(book, master, filters)
        save_state(book, master, plan)
        book.Save()
    restore_calculation_mode(path, master.calculation_mode)
    return tuple(written)
