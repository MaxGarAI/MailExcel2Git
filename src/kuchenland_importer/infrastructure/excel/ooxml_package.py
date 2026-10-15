"""Bounded read-only access to OOXML parts and internal relationships."""

import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from kuchenland_importer.infrastructure.excel.read_limits import SOURCE_LIMITS, ReadLimits

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


class OOXMLPackage:
    def __init__(self, path: Path, *, limits: ReadLimits = SOURCE_LIMITS) -> None:
        self.limits = limits
        self.archive = ZipFile(path)
        entries = self.archive.infolist()
        total = sum(e.file_size for e in entries)
        if len(entries) > limits.max_parts:
            self.archive.close()
            raise ValueError(
                f"Книга «{path.name}»: частей OOXML {len(entries)}, допустимо {limits.max_parts}."
            )
        if limits.max_uncompressed_bytes is not None and total > limits.max_uncompressed_bytes:
            self.archive.close()
            raise ValueError(
                f"Книга «{path.name}»: распакованный объём {total / 1024**2:.1f} МБ "
                f"превышает допустимые {limits.max_uncompressed_bytes / 1024**2:.0f} МБ."
            )

    def read(self, part: str, limit: int = 20 * 1024 * 1024) -> bytes:
        if self.archive.getinfo(part).file_size > limit:
            raise ValueError(f"Часть OOXML превышает лимит: {part}")
        return self.archive.read(part)

    def xml(self, part: str) -> ET.Element:
        data = self.read(part, self.limits.max_xml_bytes)
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
