"""Resolve native in-cell images via value metadata, rich values and image relationships."""

from xml.etree import ElementTree as ET

from openpyxl.utils.cell import coordinate_to_tuple

from kuchenland_importer.domain.photos import EmbeddedPhoto
from kuchenland_importer.infrastructure.excel.ooxml_package import OOXMLPackage, R, S

RD = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
RR = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"


def _at(items: ET.Element, index: int) -> ET.Element:
    if not 0 <= index < len(items):
        raise ValueError("Индекс rich data выходит за пределы списка.")
    return items[index]


def native_images(package: OOXMLPackage) -> dict[int, bytes]:
    if "xl/metadata.xml" not in package.archive.namelist():
        return {}
    metadata = package.xml("xl/metadata.xml")
    types = metadata.find(f"{{{S}}}metadataTypes")
    if types is None:
        return {}
    ids = [i for i, t in enumerate(types, 1) if t.attrib.get("name") == "XLRICHVALUE"]
    if not ids:
        return {}
    future = metadata.find(f"{{{S}}}futureMetadata[@name='XLRICHVALUE']")
    values = metadata.find(f"{{{S}}}valueMetadata")
    if future is None or values is None:
        raise ValueError("Неполная метаинформация native images.")
    data = package.xml("xl/richData/rdrichvalue.xml")
    structures = package.xml("xl/richData/rdrichvaluestructure.xml")
    rels = package.xml("xl/richData/richValueRel.xml")
    links = package.relationships("xl/richData/richValueRel.xml")
    images = {}
    for index, block in enumerate(values, 1):
        records = block.findall(f"{{{S}}}rc")
        for record in records:
            if int(record.attrib["t"]) not in ids:
                continue
            future_index = int(record.attrib["v"])
            reference = _at(future, future_index).find(f".//{{{RD}}}rvb")
            if reference is None:
                raise ValueError("Отсутствует rich value reference.")
            value = _at(data, int(reference.attrib["i"]))
            structure = _at(structures, int(value.attrib["s"]))
            if structure.attrib.get("t") != "_localImage":
                continue
            keys = [i for i, k in enumerate(structure)
                    if k.attrib.get("n") == "_rvRel:LocalImageIdentifier"]
            if len(keys) != 1:
                raise ValueError("Неоднозначный LocalImageIdentifier.")
            rel_index = int(_at(value, keys[0]).text or "-1")
            if rel_index < 0:
                raise ValueError("Недопустимый индекс native image.")
            link = _at(rels, rel_index).attrib[f"{{{R}}}id"]
            target = links[link]
            if target is None:
                raise ValueError("Внешнее native image не поддерживается.")
            images[index] = package.read(target)
    return images


def read_native(
    package: OOXMLPackage, sheet: str, part: str, images: dict[int, bytes]
) -> tuple[EmbeddedPhoto, ...]:
    result = []
    for cell in package.xml(part).findall(f"{{{S}}}sheetData/{{{S}}}row/{{{S}}}c"):
        if "vm" not in cell.attrib:
            continue
        index = int(cell.attrib["vm"])
        if index in images:
            row, column = coordinate_to_tuple(cell.attrib["r"])
            result.append(EmbeddedPhoto(
                sheet, row, column, images[index], f"{part}:{cell.attrib['r']}", native=True
            ))
    return tuple(result)
