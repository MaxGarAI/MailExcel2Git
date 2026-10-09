"""Exercise XLS/XLSB value reading and photo conversion on new test files in installed Excel."""

import argparse
import sys
from hashlib import sha256
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader  # noqa: E402
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader  # noqa: E402


def main() -> None:
    import pythoncom
    import win32com.client
    from PIL import Image

    parser = argparse.ArgumentParser(
        description="Проверка старых форматов на новых тестовых книгах"
    )
    parser.add_argument("directory", type=Path)
    root = parser.parse_args().directory.resolve()
    root.mkdir(parents=True, exist_ok=False)
    image = root / "image.png"
    Image.new("RGB", (40, 30), "blue").save(image)
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = book = None
    paths = []
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        for suffix, file_format in (("xls", 56), ("xlsb", 50)):
            book = app.Workbooks.Add()
            sheet = book.Worksheets(1)
            sheet.Range("A1:C1").Value = (("SKU", "Photo", "Name"),)
            sheet.Range("A2").Value = "QA-001"
            sheet.Range("C2").Value = "QA product"
            cell = sheet.Range("B2")
            sheet.Shapes.AddPicture(str(image), False, True, cell.Left, cell.Top, 30, 20)
            path = root / f"source.{suffix}"
            book.SaveAs(str(path), FileFormat=file_format)
            book.Close(SaveChanges=False)
            book = None
            paths.append(path)
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
    for path in paths:
        before = sha256(path.read_bytes()).digest()
        values = WorkbookReader().read(path)
        photos = PhotoReader().read(path, root / "copies")
        assert values[0].rows[1][0].value == "QA-001"
        assert len(photos.pictures) == 1 and not photos.problems
        assert (photos.pictures[0].row, photos.pictures[0].column) == (2, 2)
        assert sha256(path.read_bytes()).digest() == before
        assert not list((root / "copies").glob("photo-source-*.xlsx"))
        print(f"PASS {path.suffix}: values and photo read; source SHA-256 unchanged.")


if __name__ == "__main__":
    main()
