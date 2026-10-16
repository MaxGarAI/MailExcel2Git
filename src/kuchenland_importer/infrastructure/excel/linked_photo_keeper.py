"""Preserve external picture identity while Excel edits only data and embedded images."""

import copy
from pathlib import Path
from typing import Any
from uuid import uuid4
from xml.etree import ElementTree as ET

from kuchenland_importer.domain.write_plan import WritePlan
from kuchenland_importer.infrastructure.excel.linked_photo_package import (
    HeldPhoto,
    quarantine,
    restore_anchors,
)
from kuchenland_importer.infrastructure.excel.ooxml_package import A, D


class LinkedPhotoKeeper:
    def __init__(self, path: Path, plan: WritePlan) -> None:
        references = {
            ref
            for action in plan.products
            if action.previous
            for ref in action.previous.linked_photos
        }
        held = quarantine(path, references) if references else {}
        self.pictures = {
            action.incoming.product.article: [held[ref] for ref in action.previous.linked_photos]
            for action in plan.products
            if action.previous
            and action.previous.linked_photos
            and action.incoming.product.photo is None
        }
        self.restored: dict[str, list[HeldPhoto]] = {}

    def restore(self, article: str, sheet: Any, cell: Any) -> tuple[str, ...]:
        names = []
        for picture in self.pictures.pop(article, []):
            if cell.RowHeight < 60:
                cell.RowHeight = 100
            source = picture.anchor
            pic = copy.deepcopy(source.find(f"{{{D}}}pic"))
            if pic is None:
                raise ValueError("У прежнего внешнего фото отсутствует pic.")
            name = "KL_link_" + uuid4().hex
            properties = pic.find(f"{{{D}}}nvPicPr/{{{D}}}cNvPr")
            assert properties is not None
            properties.set("name", name)
            transform = pic.find(f"{{{D}}}spPr/{{{A}}}xfrm")
            extent = transform.find(f"{{{A}}}ext") if transform is not None else None
            width = int(extent.attrib["cx"]) if extent is not None else 1
            height = int(extent.attrib["cy"]) if extent is not None else 1
            if width <= 0 or height <= 0:
                raise ValueError("Некорректный размер прежнего внешнего фото.")
            scale = min(
                max(1, float(cell.Width) - 4) * 12700 / width,
                max(1, float(cell.Height) - 4) * 12700 / height,
            )
            width, height = max(1, round(width * scale)), max(1, round(height * scale))
            anchor = ET.Element(f"{{{D}}}twoCellAnchor", editAs="oneCell")
            for tag, x, y in (("from", 25400, 25400), ("to", 25400 + width, 25400 + height)):
                marker = ET.SubElement(anchor, f"{{{D}}}{tag}")
                for key, value in (
                    ("col", int(cell.Column) - 1),
                    ("colOff", x),
                    ("row", int(cell.Row) - 1),
                    ("rowOff", y),
                ):
                    ET.SubElement(marker, f"{{{D}}}{key}").text = str(value)
            if transform is not None:
                offset = transform.find(f"{{{A}}}off")
                if offset is not None:
                    offset.set("x", str(round((float(cell.Left) + 2) * 12700)))
                    offset.set("y", str(round((float(cell.Top) + 2) * 12700)))
                if extent is not None:
                    extent.set("cx", str(width))
                    extent.set("cy", str(height))
            anchor.append(pic)
            client = source.find(f"{{{D}}}clientData")
            anchor.append(
                copy.deepcopy(client) if client is not None else ET.Element(f"{{{D}}}clientData")
            )
            self.restored.setdefault(str(sheet.Name), []).append(
                HeldPhoto(anchor, picture.relationships)
            )
            names.append(name)
        return tuple(names)

    def finish(self, path: Path) -> None:
        if self.pictures:
            raise ValueError("Не все прежние внешние фото возвращены к товарам.")
        restore_anchors(path, self.restored)
