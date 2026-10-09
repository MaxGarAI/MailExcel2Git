import io
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from openpyxl import Workbook
from PIL import Image

from kuchenland_importer.infrastructure.excel.ooxml_package import S
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader
from kuchenland_importer.infrastructure.photo_store import PhotoStore

RD = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
RR = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def native_book(path: Path, index: int = 0, external: bool = False) -> None:
    book = Workbook()
    book.active.title = "Images"
    book.active.append(["SKU", "Photo", "Name"])
    book.active.append(["A", "#VALUE!", "First"])
    book.save(path)
    with ZipFile(path) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    sheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    sheet.find(f".//{{{S}}}c[@r='B2']").set("vm", "1")
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(sheet)
    parts["xl/metadata.xml"] = f'''<metadata xmlns="{S}" xmlns:rd="{RD}">
<metadataTypes><metadataType name="XLDAPR"/><metadataType name="XLRICHVALUE"/></metadataTypes>
<futureMetadata name="XLRICHVALUE"><bk><extLst><ext><rd:rvb i="{index}"/></ext></extLst></bk>
</futureMetadata><valueMetadata><bk><rc t="2" v="0"/></bk></valueMetadata></metadata>'''.encode()
    parts["xl/richData/rdrichvalue.xml"] = (
        f'<rvData xmlns="{RD}"><rv s="0"><v>5</v><v>0</v></rv></rvData>'.encode()
    )
    parts["xl/richData/rdrichvaluestructure.xml"] = f'''<rvStructures xmlns="{RD}">
<s t="_localImage"><k n="CalcOrigin"/><k n="_rvRel:LocalImageIdentifier"/></s>
</rvStructures>'''.encode()
    parts["xl/richData/richValueRel.xml"] = (
        f'<richValueRels xmlns="{RR}" xmlns:r="{R}"><rel r:id="custom"/></richValueRels>'.encode()
    )
    mode = 'TargetMode="External"' if external else ''
    parts["xl/richData/_rels/richValueRel.xml.rels"] = (
        f'<Relationships xmlns="{REL}"><Relationship Id="custom" '
        f'Target="../media/native.png" {mode}/></Relationships>'.encode()
    )
    stream = io.BytesIO()
    Image.new("RGB", (50, 30), "blue").save(stream, format="PNG")
    parts["xl/media/native.png"] = stream.getvalue()
    with ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def test_native_metadata_types_and_key_order_are_resolved(tmp_path: Path) -> None:
    path = tmp_path / "native.xlsx"
    native_book(path)
    before = path.read_bytes()
    result = PhotoReader().read(path)
    assert len(result.pictures) == 1 and not result.problems
    photo = result.pictures[0]
    assert (photo.sheet, photo.row, photo.column, photo.native) == ("Images", 2, 2, True)
    asset = PhotoStore(tmp_path / "images").build(result.pictures)
    assert (asset.width, asset.height) == (50, 30)
    assert path.read_bytes() == before


@pytest.mark.parametrize("index", [-1, 1])
def test_invalid_rich_value_index_is_rejected(tmp_path: Path, index: int) -> None:
    path = tmp_path / "bad-native.xlsx"
    native_book(path, index=index)
    with pytest.raises(ValueError, match="Индекс"):
        PhotoReader().read(path)


def test_external_native_image_is_not_downloaded(tmp_path: Path) -> None:
    path = tmp_path / "external.xlsx"
    native_book(path, external=True)
    with pytest.raises(ValueError, match="Внешнее"):
        PhotoReader().read(path)
