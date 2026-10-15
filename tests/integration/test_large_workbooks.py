import io
import random
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook
from openpyxl.drawing.image import Image as DrawingImage
from PIL import Image

from kuchenland_importer.domain.season import SeasonRoute
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.master_snapshot import read_snapshot
from kuchenland_importer.infrastructure.excel.ooxml_package import OOXMLPackage
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader
from kuchenland_importer.infrastructure.excel.read_limits import MASTER_LIMITS, ReadLimits
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader


def test_real_workbook_above_50_mib_reads_values_and_all_photos(tmp_path):
    path = tmp_path / "large-master.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "ФОТО весна"
    sheet.append(["Артикул", "Поставщик", "Сезон", "Наименование", "Фото"])
    stream = io.BytesIO()
    Image.frombytes("RGB", (1024, 1024), random.Random(42).randbytes(1024 * 1024 * 3)).save(
        stream, format="PNG"
    )
    for index in range(17):
        row = index + 2
        sheet.append([f"SKU-{index:03}", "Vendor", "Весна", "Product"])
        sheet.row_dimensions[row].height = 100
        image = DrawingImage(io.BytesIO(stream.getvalue()))
        image.width = image.height = 80
        sheet.add_image(image, f"E{row}")
    book.save(path)
    book.close()
    assert path.stat().st_size > 50 * 1024 * 1024
    assert len(WorkbookReader().read(path)[0].rows) == 18
    assert len(PhotoReader().read(path).pictures) == 17
    snapshot = read_snapshot(
        path,
        load_columns(Path("config/columns.toml")),
        (SeasonRoute("spring", ("Весна",), "Весна", "ФОТО весна"),),
        (),
        tmp_path / "photos",
    )
    assert len(snapshot.products) == 17
    assert all(product.photo is not None for product in snapshot.products.values())


def test_master_has_no_attachment_aggregate_budget_but_source_still_does(tmp_path):
    path = tmp_path / "compressed.xlsx"
    book = Workbook()
    book.active.append(["SKU", "Supplier", "Season"])
    book.active.append(["ONE", "Vendor", "Весна"])
    book.save(path)
    book.close()
    with ZipFile(path, "a", compression=ZIP_DEFLATED) as package:
        with package.open("unused.bin", "w") as stream:
            for _ in range(257):
                stream.write(b"0" * 1024 * 1024)
    with pytest.raises(ValueError, match="распакованный объём"):
        WorkbookReader().read(path)
    assert WorkbookReader().read(path, limits=MASTER_LIMITS)[0].rows[1][0].value == "ONE"
    assert not PhotoReader().read(path, limits=MASTER_LIMITS).pictures


def test_structural_parts_budget_is_still_enforced(tmp_path):
    path = tmp_path / "parts.xlsx"
    with ZipFile(path, "w") as archive:
        archive.writestr("one", b"1")
        archive.writestr("two", b"2")
    with pytest.raises(ValueError, match="частей OOXML 2"):
        OOXMLPackage(path, limits=ReadLimits(max_parts=1))
