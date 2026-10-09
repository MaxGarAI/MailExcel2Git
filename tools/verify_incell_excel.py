"""Acceptance check on a new disposable book; never opens the master workbook."""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from kuchenland_importer.domain.photos import EmbeddedPhoto  # noqa: E402
from kuchenland_importer.infrastructure.excel.incell_writer import InCellPictureWriter  # noqa: E402
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader  # noqa: E402
from kuchenland_importer.infrastructure.photo_store import PhotoStore  # noqa: E402


def main() -> None:
    import pythoncom
    import win32com.client
    from PIL import Image

    parser = argparse.ArgumentParser(description="Проверка нативных фото на новой тестовой книге")
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    root = args.directory.resolve()
    root.mkdir(parents=True, exist_ok=False)
    image = root / "input.png"
    Image.new("RGB", (90, 60), (240, 80, 40)).save(image)
    store = PhotoStore(root / "photos")
    asset = store.build((EmbeddedPhoto("Sheet1", 2, 2, image.read_bytes(), "QA"),))
    output = root / "native.xlsx"
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = book = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        book = app.Workbooks.Add()
        sheet = book.Worksheets(1)
        sheet.Name = "Photos"
        sheet.Range("A1:B1").Value = (("SKU", "Photo"),)
        sheet.Range("A2").Value = "QA-001"
        sheet.Range("B2").RowHeight = 80
        sheet.Range("B2").ColumnWidth = 18
        InCellPictureWriter().place(sheet, "B2", asset)
        book.SaveAs(str(output), FileFormat=51)
        book.Close(SaveChanges=False)
        book = app.Workbooks.Open(str(output), ReadOnly=True, UpdateLinks=0)
        assert book.Worksheets(1).Shapes.Count == 0
        pictures = PhotoReader().read(output).pictures
        assert len(pictures) == 1 and pictures[0].native
        assert (pictures[0].sheet, pictures[0].row, pictures[0].column) == ("Photos", 2, 2)
        assert store.build(pictures).pixel_sha256 == asset.pixel_sha256
        print(f"PASS: Excel {app.Version}, build {app.Build}; native image survives save/reopen.")
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


if __name__ == "__main__":
    main()
