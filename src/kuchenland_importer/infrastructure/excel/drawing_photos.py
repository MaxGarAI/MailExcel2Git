"""Read floating pictures, including their visible crop and orientation."""

from typing import cast
from xml.etree import ElementTree as ET

from kuchenland_importer.domain.photos import EmbeddedPhoto, PhotoExtraction, PhotoProblem
from kuchenland_importer.infrastructure.excel.drawing_geometry import RowGeometry
from kuchenland_importer.infrastructure.excel.ooxml_package import A, D, OOXMLPackage, R, S


def one_cell_end(sheet: ET.Element, row: int, offset: int, height: int) -> int:
    return RowGeometry.from_sheet(sheet).extent_bounds(row, offset, height)[1]


def read_drawings(package: OOXMLPackage, sheet: str, part: str) -> PhotoExtraction:
    pictures: list[EmbeddedPhoto] = []
    problems: list[PhotoProblem] = []
    links = package.relationships(part)
    worksheet = package.xml(part)
    geometry = RowGeometry.from_sheet(worksheet)
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
                    if (
                        anchor.find(f"{{{D}}}pic") is not None
                        or anchor.find(f"{{{D}}}grpSp") is not None
                    ):
                        raise ValueError("У картинки нет привязки к строке/столбцу.")
                    continue
                row = int(start.findtext(f"{{{D}}}row", "-1")) + 1
                row, row_offset = geometry.normalize(
                    row, int(start.findtext(f"{{{D}}}rowOff", "0"))
                )
                column = int(start.findtext(f"{{{D}}}col", "-1")) + 1
                pic = anchor.find(f"{{{D}}}pic")
                if pic is None:
                    if anchor.find(f"{{{D}}}grpSp") is not None:
                        raise ValueError("Группированные картинки требуют отдельного профиля.")
                    continue
                blip = pic.find(f"{{{D}}}blipFill/{{{A}}}blip")
                if blip is not None and f"{{{R}}}embed" not in blip.attrib:
                    link = blip.attrib.get(f"{{{R}}}link")
                    properties = pic.find(f"{{{D}}}nvPicPr/{{{D}}}cNvPr")
                    name = properties.attrib.get("name") if properties is not None else None
                    if link in media and media[link] is None and name:
                        problems.append(PhotoProblem(
                            sheet, row, column,
                            "Фото хранится только во внешней ссылке, без картинки внутри книги. "
                            "Данные импортируются без загрузки фото из сети; "
                            "прежнее фото сохраняется.",
                            external_link=True, shape_name=name,
                        ))
                        continue
                if blip is None or f"{{{R}}}embed" not in blip.attrib:
                    raise ValueError(
                        "Отсутствует встроенная картинка или корректная внешняя ссылка."
                    )
                target = media[blip.attrib[f"{{{R}}}embed"]]
                if target is None:
                    raise ValueError("Внешнее изображение не поддерживается.")
                crop = pic.find(f"{{{D}}}blipFill/{{{A}}}srcRect")
                transform = pic.find(f"{{{D}}}spPr/{{{A}}}xfrm")
                end = anchor.find(f"{{{D}}}to")
                extent = anchor.find(f"{{{D}}}ext")
                transform_extent = (
                    transform.find(f"{{{A}}}ext") if transform is not None else None
                )
                zero_extent = (
                    int(extent.attrib["cy"]) == 0 if extent is not None else
                    transform_extent is not None and int(transform_extent.attrib["cy"]) == 0
                )
                zero_anchor = end is None or geometry.normalize(
                    int(end.findtext(f"{{{D}}}row", "-1")) + 1,
                    int(end.findtext(f"{{{D}}}rowOff", "0")),
                ) == (row, row_offset)
                if zero_extent and zero_anchor:
                    problems.append(PhotoProblem(
                        sheet, row, column,
                        "Невидимый объект фото с нулевой высотой пропущен.",
                        non_rendering=True,
                    ))
                    continue
                if end is not None:
                    row, end_row = geometry.bounds(
                        row,
                        row_offset,
                        int(end.findtext(f"{{{D}}}row", "-1")) + 1,
                        int(end.findtext(f"{{{D}}}rowOff", "0")),
                    )
                else:
                    row, end_row = geometry.extent_bounds(
                        row,
                        row_offset,
                        int(extent.attrib["cy"]) if extent is not None else 0,
                    )
                pictures.append(
                    EmbeddedPhoto(
                        sheet,
                        row,
                        column,
                        package.read(target),
                        f"{drawing_part}:{index}",
                        max(row, end_row),
                        cast(
                            tuple[int, int, int, int],
                            tuple(int(crop.attrib.get(k, "0")) for k in ("l", "t", "r", "b")),
                        )
                        if crop is not None
                        else (0, 0, 0, 0),
                        float(transform.attrib.get("rot", "0")) / 60000
                        if transform is not None
                        else 0,
                        transform is not None and transform.attrib.get("flipH") in {"1", "true"},
                        transform is not None and transform.attrib.get("flipV") in {"1", "true"},
                    )
                )
            except (ValueError, KeyError) as error:
                problems.append(PhotoProblem(sheet, row, column, str(error)))
    return PhotoExtraction(tuple(pictures), tuple(problems))
