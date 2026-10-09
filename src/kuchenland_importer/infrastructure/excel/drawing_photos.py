"""Read floating pictures, including their visible crop and orientation."""

from typing import cast
from xml.etree import ElementTree as ET

from kuchenland_importer.domain.photos import EmbeddedPhoto, PhotoExtraction, PhotoProblem
from kuchenland_importer.infrastructure.excel.ooxml_package import A, D, OOXMLPackage, R, S


def one_cell_end(sheet: ET.Element, row: int, offset: int, height: int) -> int:
    formatting = sheet.find(f"{{{S}}}sheetFormatPr")
    default = (
        float(formatting.attrib.get("defaultRowHeight", "15"))
        if formatting is not None else 15
    )
    heights = {
        int(r.attrib["r"]): 0 if r.attrib.get("hidden") in {"1", "true"}
        else float(r.attrib.get("ht", str(default)))
        for r in sheet.findall(f"{{{S}}}sheetData/{{{S}}}row")
    }
    remaining = offset + height
    if offset < 0 or height <= 0:
        raise ValueError("Недопустимый размер/смещение фото.")
    for end in range(row, 100001):
        # Excel row heights use points; DrawingML uses 12700 EMU per point.
        remaining -= round(heights.get(end, default) * 12700)
        if remaining <= 0:
            return end
    raise ValueError("Фото выходит за пределы поддерживаемых строк.")


def read_drawings(package: OOXMLPackage, sheet: str, part: str) -> PhotoExtraction:
    pictures: list[EmbeddedPhoto] = []
    problems: list[PhotoProblem] = []
    links = package.relationships(part)
    worksheet = package.xml(part)
    for drawing in worksheet.findall(f"{{{S}}}drawing"):
        drawing_part = links[drawing.attrib[f"{{{R}}}id"]]
        if drawing_part is None:
            raise ValueError("Внешний drawing не поддерживается.")
        media = package.relationships(drawing_part)
        for index, anchor in enumerate(package.xml(drawing_part), 1):
            row = column = None
            try:
                start = anchor.find(f"{{{D}}}from")
                if start is None:
                    if (anchor.find(f"{{{D}}}pic") is not None
                            or anchor.find(f"{{{D}}}grpSp") is not None):
                        raise ValueError("У картинки нет привязки к строке/столбцу.")
                    continue
                row = int(start.findtext(f"{{{D}}}row", "-1")) + 1
                column = int(start.findtext(f"{{{D}}}col", "-1")) + 1
                pic = anchor.find(f"{{{D}}}pic")
                if pic is None:
                    if anchor.find(f"{{{D}}}grpSp") is not None:
                        raise ValueError("Группированные картинки требуют отдельного профиля.")
                    continue
                blip = pic.find(f"{{{D}}}blipFill/{{{A}}}blip")
                if blip is None or f"{{{R}}}embed" not in blip.attrib:
                    raise ValueError("Связанная/отсутствующая картинка не загружается из сети.")
                target = media[blip.attrib[f"{{{R}}}embed"]]
                if target is None:
                    raise ValueError("Внешнее изображение не поддерживается.")
                crop = pic.find(f"{{{D}}}blipFill/{{{A}}}srcRect")
                transform = pic.find(f"{{{D}}}spPr/{{{A}}}xfrm")
                end = anchor.find(f"{{{D}}}to")
                extent = anchor.find(f"{{{D}}}ext")
                end_row = (
                    int(end.findtext(f"{{{D}}}row", "-1"))
                    + (int(end.findtext(f"{{{D}}}rowOff", "0")) > 0)
                    if end is not None else one_cell_end(
                        worksheet, row, int(start.findtext(f"{{{D}}}rowOff", "0")),
                        int(extent.attrib["cy"]) if extent is not None else 0,
                    )
                )
                pictures.append(EmbeddedPhoto(
                    sheet, row, column, package.read(target), f"{drawing_part}:{index}",
                    max(row, end_row),
                    cast(tuple[int, int, int, int],
                         tuple(int(crop.attrib.get(k, "0")) for k in ("l", "t", "r", "b")))
                    if crop is not None else (0, 0, 0, 0),
                    float(transform.attrib.get("rot", "0")) / 60000
                    if transform is not None else 0,
                    transform is not None and transform.attrib.get("flipH") in {"1", "true"},
                    transform is not None and transform.attrib.get("flipV") in {"1", "true"},
                ))
            except (ValueError, KeyError) as error:
                problems.append(PhotoProblem(sheet, row, column, str(error)))
    return PhotoExtraction(tuple(pictures), tuple(problems))
