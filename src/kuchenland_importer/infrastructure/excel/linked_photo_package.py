"""Quarantine external-only anchors before Excel opens a copy; restore without network I/O."""

import copy
import os
import posixpath
import re
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from kuchenland_importer.domain.master import LinkedPhotoReference
from kuchenland_importer.infrastructure.excel.ooxml_package import D, OOXMLPackage, R, S
from kuchenland_importer.infrastructure.excel.read_limits import MASTER_LIMITS

P = "http://schemas.openxmlformats.org/package/2006/relationships"
ANCHOR = re.compile(
    rb"<(?P<tag>(?:[\w.-]+:)?(?:twoCellAnchor|oneCellAnchor))\b.*?</(?P=tag)>", re.S
)


@dataclass(frozen=True, slots=True)
class HeldPhoto:
    anchor: ET.Element
    relationships: tuple[ET.Element, ...]


def patch_package(path: Path, updates: dict[str, bytes]) -> None:
    partial = path.with_name(".kl-linked-" + uuid4().hex + ".zip")
    try:
        with ZipFile(path) as source, ZipFile(partial, "x") as target:
            target.comment = source.comment
            for entry in source.infolist():
                target.writestr(entry, updates.get(entry.filename, source.read(entry.filename)))
            for name in updates.keys() - set(source.namelist()):
                target.writestr(name, updates[name])
        os.replace(partial, path)
    finally:
        partial.unlink(missing_ok=True)


def relpart(part: str) -> str:
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def append_xml(data: bytes, fragment: bytes) -> bytes:
    closings = list(re.finditer(rb"</[\w:.-]+\s*>", data))
    if not closings:
        root = re.search(rb"<([\w:.-]+)\b[^>]*?/>", data)
        if root is None:
            raise ValueError("Не найден корневой XML-элемент.")
        return (
            data[: root.end() - 2] + b">" + fragment + b"</" + root[1] + b">" + data[root.end() :]
        )
    closing = closings[-1]
    return data[: closing.start()] + fragment + data[closing.start() :]


def quarantine(
    path: Path, references: set[LinkedPhotoReference]
) -> dict[LinkedPhotoReference, HeldPhoto]:
    held = {}
    updates = {}
    package = OOXMLPackage(path, limits=MASTER_LIMITS)
    try:
        for sheet, part in package.sheets():
            selected = {r.shape_name: r for r in references if r.sheet == sheet}
            if not selected:
                continue
            links = package.relationships(part)
            for drawing in package.xml(part).findall(f"{{{S}}}drawing"):
                drawing_part = links[drawing.attrib[f"{{{R}}}id"]]
                if drawing_part is None:
                    raise ValueError("Внешний drawing не поддерживается.")
                raw = package.read(drawing_part)
                matches = list(ANCHOR.finditer(raw))
                parsed = [
                    a
                    for a in package.xml(drawing_part)
                    if a.tag in {f"{{{D}}}twoCellAnchor", f"{{{D}}}oneCellAnchor"}
                ]
                if len(matches) != len(parsed):
                    raise ValueError("Неоднозначная структура drawing для сохранения ссылки.")
                rels = package.xml(relpart(drawing_part))
                removed = set()
                spans = []
                for match, anchor in zip(matches, parsed, strict=True):
                    prop = anchor.find(f"{{{D}}}pic/{{{D}}}nvPicPr/{{{D}}}cNvPr")
                    name = prop.attrib.get("name") if prop is not None else None
                    if name not in selected:
                        continue
                    ids = {
                        v
                        for e in anchor.iter()
                        for k, v in e.attrib.items()
                        if k.startswith("{" + R + "}")
                    }
                    records = tuple(copy.deepcopy(r) for r in rels if r.attrib["Id"] in ids)
                    if len(records) != len(ids) or any(
                        r.attrib.get("TargetMode") != "External" for r in records
                    ):
                        raise ValueError(
                            "У внешнего фото есть неподдерживаемая внутренняя зависимость."
                        )
                    held[selected[name]] = HeldPhoto(copy.deepcopy(anchor), records)
                    removed.update(ids)
                    spans.append((match.start(), match.end()))
                if spans:
                    for start, end in reversed(spans):
                        raw = raw[:start] + raw[end:]
                    remaining = package.xml(drawing_part)
                    for anchor in list(remaining):
                        prop = anchor.find(f"{{{D}}}pic/{{{D}}}nvPicPr/{{{D}}}cNvPr")
                        if prop is not None and prop.attrib.get("name") in selected:
                            remaining.remove(anchor)
                    used = {
                        v
                        for e in remaining.iter()
                        for k, v in e.attrib.items()
                        if k.startswith("{" + R + "}")
                    }
                    for relationship in list(rels):
                        if relationship.attrib["Id"] in removed - used:
                            rels.remove(relationship)
                    updates[drawing_part] = raw
                    updates[relpart(drawing_part)] = ET.tostring(rels)
        if held.keys() != references:
            raise ValueError("Не все прежние внешние фото найдены в рабочей копии.")
    finally:
        package.close()
    if updates:
        patch_package(path, updates)
    return held


def restore_anchors(path: Path, pictures: dict[str, list[HeldPhoto]]) -> None:
    if not pictures:
        return
    updates = {}
    package = OOXMLPackage(path, limits=MASTER_LIMITS)
    try:
        for name, part in package.sheets():
            if name not in pictures:
                continue
            worksheet = package.xml(part)
            drawing = worksheet.find(f"{{{S}}}drawing")
            if drawing is None:
                drawing_part = "xl/drawings/kl-linked-" + uuid4().hex + ".xml"
                raw = f'<xdr:wsDr xmlns:xdr="{D}"/>'.encode()
                raw = raw.replace(b"/>", b"></xdr:wsDr>")
                rels = ET.Element(f"{{{P}}}Relationships")
                sheet_rels = (
                    package.xml(relpart(part))
                    if relpart(part) in package.archive.namelist()
                    else ET.Element(f"{{{P}}}Relationships")
                )
                rid = "rIdKL" + uuid4().hex
                ET.SubElement(
                    sheet_rels,
                    f"{{{P}}}Relationship",
                    Id=rid,
                    Type=R + "/drawing",
                    Target="/" + drawing_part,
                )
                updates[relpart(part)] = ET.tostring(sheet_rels)
                fragment = f'<drawing xmlns="{S}" xmlns:r="{R}" r:id="{rid}"/>'.encode()
                sheet_raw = package.read(part)
                tail = re.search(
                    rb"<(?:[\w.-]+:)?(?:legacyDrawing|legacyDrawingHF|picture|"
                    rb"oleObjects|controls|webPublishItems|tableParts|extLst)\b",
                    sheet_raw,
                )
                updates[part] = (
                    sheet_raw[: tail.start()] + fragment + sheet_raw[tail.start() :]
                    if tail
                    else append_xml(sheet_raw, fragment)
                )
                types = updates.get("[Content_Types].xml", package.read("[Content_Types].xml"))
                updates["[Content_Types].xml"] = append_xml(
                    types,
                    (
                        "<Override xmlns="
                        '"http://schemas.openxmlformats.org/package/2006/content-types" '
                        f'PartName="/{drawing_part}" ContentType="application/vnd.openxmlformats-'
                        'officedocument.drawing+xml"/>'
                    ).encode(),
                )
            else:
                resolved = package.relationships(part)[drawing.attrib[f"{{{R}}}id"]]
                if resolved is None:
                    raise ValueError("Внешний drawing назначения не поддерживается.")
                drawing_part = resolved
                raw = package.read(drawing_part)
                rels = (
                    package.xml(relpart(drawing_part))
                    if relpart(drawing_part) in package.archive.namelist()
                    else ET.Element(f"{{{P}}}Relationships")
                )
            next_id = (
                max(
                    (
                        int(p.attrib.get("id", "0"))
                        for p in ET.fromstring(raw).iter()
                        if p.tag == f"{{{D}}}cNvPr"
                    ),
                    default=0,
                )
                + 1
            )
            for picture in pictures[name]:
                anchor = copy.deepcopy(picture.anchor)
                properties = anchor.find(f"{{{D}}}pic/{{{D}}}nvPicPr/{{{D}}}cNvPr")
                assert properties is not None
                properties.set("id", str(next_id))
                next_id += 1
                remap = {}
                for record in picture.relationships:
                    rel = copy.deepcopy(record)
                    rid = "rIdKL" + uuid4().hex
                    remap[rel.attrib["Id"]] = rid
                    rel.set("Id", rid)
                    rels.append(rel)
                for element in anchor.iter():
                    for key, value in list(element.attrib.items()):
                        if key.startswith("{" + R + "}"):
                            element.set(key, remap[value])
                raw = append_xml(raw, ET.tostring(anchor))
            updates[drawing_part] = raw
            updates[relpart(drawing_part)] = ET.tostring(rels)
    finally:
        package.close()
    patch_package(path, updates)
