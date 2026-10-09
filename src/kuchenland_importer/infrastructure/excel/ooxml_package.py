"""Bounded read-only access to OOXML parts and internal relationships."""

import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


class OOXMLPackage:
    def __init__(self, path: Path) -> None:
        if path.stat().st_size > 50 * 1024 * 1024:
            raise ValueError("Книга превышает лимит чтения 50 МБ.")
        self.archive = ZipFile(path)
        entries = self.archive.infolist()
        if len(entries) > 10000 or sum(e.file_size for e in entries) > 256 * 1024 * 1024:
            self.archive.close()
            raise ValueError("OOXML превышает лимит распакованного размера / числа частей.")

    def read(self, part: str, limit: int = 20 * 1024 * 1024) -> bytes:
        if self.archive.getinfo(part).file_size > limit:
            raise ValueError(f"Часть OOXML превышает лимит: {part}")
        return self.archive.read(part)

    def xml(self, part: str) -> ET.Element:
        data = self.read(part, 16 * 1024 * 1024)
        if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
            raise ValueError("Объявления DTD/ENTITY в OOXML не поддерживаются.")
        return ET.fromstring(data)

    def relationships(self, part: str) -> dict[str, str | None]:
        folder, name = posixpath.split(part)
        relpart = posixpath.join(folder, "_rels", name + ".rels")
        if relpart not in self.archive.namelist():
            return {}
        result: dict[str, str | None] = {}
        for rel in self.xml(relpart):
            target = rel.attrib["Target"]
            if rel.attrib.get("TargetMode") == "External":
                result[rel.attrib["Id"]] = None
                continue
            resolved = posixpath.normpath(
                target.lstrip("/") if target.startswith("/") else posixpath.join(folder, target)
            )
            if resolved.startswith("../") or ":" in resolved or "\\" in resolved:
                raise ValueError("Недопустимый путь части OOXML.")
            result[rel.attrib["Id"]] = resolved
        return result

    def sheets(self) -> tuple[tuple[str, str], ...]:
        relationships = self.relationships("xl/workbook.xml")
        result = []
        for sheet in self.xml("xl/workbook.xml").findall(f"{{{S}}}sheets/{{{S}}}sheet"):
            part = relationships[sheet.attrib[f"{{{R}}}id"]]
            if part is not None:
                result.append((sheet.attrib["name"], part))
        return tuple(result)

    def close(self) -> None:
        self.archive.close()
