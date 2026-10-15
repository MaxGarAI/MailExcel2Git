from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook
from openpyxl.drawing.image import Image as DrawingImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from PIL import Image

from kuchenland_importer.infrastructure.excel.drawing_geometry import RowGeometry
from kuchenland_importer.infrastructure.excel.ooxml_package import S
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader


def test_offset_carried_to_real_row_does_not_bind_picture_to_previous_article(tmp_path):
    image = tmp_path / "blue.png"
    Image.new("RGB", (32, 32), "blue").save(image)
    path = tmp_path / "offset.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Фото"
    sheet.append(["Article", "Photo", "Name"])
    sheet.append(["A", None, "First"])
    sheet.append(["B", None, "Second"])
    for row in (2, 3):
        sheet.row_dimensions[row].height = 100
    picture = DrawingImage(image)
    picture.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=1, row=1, rowOff=1_270_000 + 100),
        to=AnchorMarker(col=1, row=2, rowOff=1_200_000),
    )
    sheet.add_image(picture)
    book.save(path)
    book.close()
    photo = PhotoReader().read(path).pictures[0]
    assert photo.row == photo.end_row == 3


def test_real_cross_row_picture_still_spans_both_rows(tmp_path):
    image = tmp_path / "red.png"
    Image.new("RGB", (32, 32), "red").save(image)
    book = Workbook()
    book.active.append(["Article", "Photo", "Name"])
    book.active.append(["A", None, "First"])
    book.active.append(["B", None, "Second"])
    book.active.row_dimensions[2].height = 100
    book.active.row_dimensions[3].height = 100
    picture = DrawingImage(image)
    picture.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=1, row=1, rowOff=100), to=AnchorMarker(col=1, row=2, rowOff=500000)
    )
    book.active.add_image(picture)
    path = tmp_path / "spans.xlsx"
    book.save(path)
    book.close()
    photo = PhotoReader().read(path).pictures[0]
    assert (photo.row, photo.end_row) == (2, 3)


def test_offset_can_cross_multiple_rows_and_exact_boundary():
    geometry = RowGeometry.from_sheet(
        ET.fromstring(f'<worksheet xmlns="{S}"><sheetFormatPr defaultRowHeight="100"/></worksheet>')
    )
    assert geometry.normalize(2, 1_270_000) == (3, 0)
    assert geometry.normalize(2, 2_540_010) == (4, 10)
    with pytest.raises(ValueError):
        geometry.normalize(2, -1)


def test_hidden_row_keeps_its_logical_image_anchor():
    geometry = RowGeometry.from_sheet(
        ET.fromstring(
            f'<worksheet xmlns="{S}"><sheetFormatPr defaultRowHeight="15"/>'
            '<sheetData><row r="2" ht="267" hidden="1"/></sheetData></worksheet>'
        )
    )
    assert geometry.normalize(2, 500000) == (2, 500000)
    assert geometry.bounds(2, 100000, 2, 2000000) == (2, 2)


def test_tiny_border_overhang_is_not_a_second_product_but_material_overlap_is():
    geometry = RowGeometry.from_sheet(
        ET.fromstring(
            f'<worksheet xmlns="{S}"><sheetFormatPr defaultRowHeight="166.05"/></worksheet>'
        )
    )
    assert geometry.bounds(3, 38100, 3, 2148640) == (3, 3)
    assert geometry.bounds(3, 38100, 4, 500000) == (3, 4)


def test_one_cell_extent_uses_hidden_row_logical_height_and_offset_carry():
    geometry = RowGeometry.from_sheet(
        ET.fromstring(
            f'<worksheet xmlns="{S}"><sheetFormatPr defaultRowHeight="100"/>'
            '<sheetData><row r="2" ht="100" hidden="1"/></sheetData></worksheet>'
        )
    )
    assert geometry.extent_bounds(2, 100, 1000000) == (2, 2)
    assert geometry.extent_bounds(2, 1270100, 1000000) == (3, 3)
