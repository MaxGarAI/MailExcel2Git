"""Disposable native Excel acceptance for filter criteria after a range is extended."""

import argparse
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pythoncom  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from win32com.client import VARIANT  # noqa: E402

from kuchenland_importer.domain.master import MasterSnapshot, SheetLayout  # noqa: E402
from kuchenland_importer.infrastructure.excel.master_filters import (  # noqa: E402
    capture_filters,
    restore_filters,
    suspend_filters,
)
from kuchenland_importer.infrastructure.excel.owned_session import owned_workbook  # noqa: E402


def fixture(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "Products"
    sheet.append(["Article", "Color", "Date", "Price", "Note"])
    for number in range(1, 10):
        sheet.append(
            [
                str(number),
                ("Red", "Blue", "Green")[number % 3],
                datetime(2026, 10 if number % 2 else 9, number),
                number * 10,
                None if number % 2 else "Note",
            ]
        )
    sheet.auto_filter.ref = "A1:E10"
    book.save(path)
    book.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    root = parser.parse_args().directory.resolve()
    root.mkdir(parents=True, exist_ok=False)
    path = root / "filters.xlsx"
    fixture(path)
    layout = SheetLayout("Products", 1, 1, 2, {"article": 1}, 10)
    snapshot = MasterSnapshot({"Products": layout}, {}, ("Products",))
    missing = VARIANT(pythoncom.VT_ERROR, -2147352572)
    cases = (
        ("single", 2, "Red", missing, missing),
        ("blank", 5, "=", missing, missing),
        ("nonblank", 5, "<>", missing, missing),
        ("multi", 2, ("Red", "Blue", "Green"), 7, missing),
        ("and", 4, ">=20", 1, "<=80"),
        ("or", 2, "Red", 2, "Blue"),
        ("date_range", 3, ">=10/1/2026", 1, "<=10/31/2026"),
        ("dynamic", 3, 8, 11, missing),
        ("cell_color", 2, 255, 8, missing),
        ("font_color", 2, 255, 9, missing),
        ("top_items", 4, "3", 3, missing),
    )
    with owned_workbook(path) as book:
        sheet = book.Worksheets(1)
        sheet.Range("B2:B4").Interior.Color = 255
        sheet.Range("B2:B4").Font.Color = 255
        area = sheet.Range("A1:E10")
        for name, field, first, operator, second in cases:
            if sheet.FilterMode:
                sheet.ShowAllData()
            sheet.AutoFilterMode = False
            # Set up criteria independently of the production helper.
            dispid = area._oleobj_.GetIDsOfNames("AutoFilter")
            area._oleobj_.Invoke(dispid, 0, 1, True, field, first, operator, second)
            book.Save()
            before = capture_filters(book, snapshot)
            visible = [not sheet.Rows(r).Hidden for r in range(2, 11)]
            suspend_filters(book, before)
            assert not sheet.FilterMode
            assert sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row == 10
            sheet.Cells(11, 1).Value = "NEW"
            sheet.Cells(11, 2).Value = "Red"
            restore_filters(book, snapshot, before)
            after = capture_filters(book, snapshot)
            assert before == after, (name, before, after)
            assert visible == [not sheet.Rows(r).Hidden for r in range(2, 11)], name
            assert sheet.AutoFilter.Range.Rows.Count == 11
            sheet.Rows(11).Delete()
            print(name + ": PASS", flush=True)
        book.Save()
    print("PASS: 11 filter types, criteria, row visibility and range extension.")


if __name__ == "__main__":
    main()
