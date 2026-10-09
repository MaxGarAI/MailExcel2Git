# Этап 5: полные файлы для ревью

Версия 0.5.0. Полные новые и изменённые файлы; рабочие письма, книги и фото не включены.

## src/kuchenland_importer/domain/photos.py

````python
"""Embedded image payloads and failures with source cell coordinates."""

from dataclasses import dataclass
from enum import StrEnum

from kuchenland_importer.domain.product import PhotoAsset


@dataclass(frozen=True, slots=True)
class EmbeddedPhoto:
    sheet: str
    row: int
    column: int
    data: bytes
    source_id: str
    end_row: int | None = None
    crop: tuple[int, int, int, int] = (0, 0, 0, 0)
    rotation: float = 0
    flip_h: bool = False
    flip_v: bool = False
    native: bool = False

    def __post_init__(self) -> None:
        if not (1 <= self.row <= 100000 and 1 <= self.column <= 512):
            raise ValueError("Фото находится за пределами поддерживаемых координат.")
        if self.end_row is not None and not self.row <= self.end_row <= 100000:
            raise ValueError("Недопустимая конечная строка фото.")
        if not self.data or len(self.data) > 20 * 1024 * 1024:
            raise ValueError("Пустое фото или превышен лимит 20 МБ.")


@dataclass(frozen=True, slots=True)
class PhotoProblem:
    sheet: str
    row: int | None
    column: int | None
    message: str


@dataclass(frozen=True, slots=True)
class PhotoExtraction:
    pictures: tuple[EmbeddedPhoto, ...]
    problems: tuple[PhotoProblem, ...] = ()


class PhotoAction(StrEnum):
    ADD = "add"
    REPLACE = "replace"
    KEEP = "keep"
    NONE = "none"


def photo_action(existing: PhotoAsset | None, incoming: PhotoAsset | None) -> PhotoAction:
    if incoming is None:
        return PhotoAction.KEEP if existing is not None else PhotoAction.NONE
    if existing is None:
        return PhotoAction.ADD
    return (
        PhotoAction.KEEP
        if existing.pixel_sha256 == incoming.pixel_sha256
        else PhotoAction.REPLACE
    )
````

## src/kuchenland_importer/domain/product.py

````python
"""A product row and its provenance; source formulas are not part of this model."""

import math
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from types import MappingProxyType

from kuchenland_importer.domain.cell_values import CellValue, ExcelErrorValue
from kuchenland_importer.domain.mail import SavedAttachment


@dataclass(frozen=True, slots=True)
class PhotoAsset:
    path: Path
    pixel_sha256: str
    width: int
    height: int
    image_count: int = 1

    def __post_init__(self) -> None:
        if not self.path.is_absolute() or self.width < 1 or self.height < 1:
            raise ValueError("Фотография должна иметь абсолютный путь и положительный размер.")
        if type(self.image_count) is not int or not 1 <= self.image_count <= 16:
            raise ValueError("Фото должно содержать от 1 до 16 разных изображений.")
        digest = self.pixel_sha256
        if len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest):
            raise ValueError("Ожидается SHA-256 нормализованных пикселей.")


@dataclass(frozen=True, slots=True)
class ProductRecord:
    article: str
    supplier: str
    season: str
    attachment: SavedAttachment
    source_sheet: str
    source_row: int
    values: Mapping[str, CellValue]
    photo: PhotoAsset | None = None

    def __post_init__(self) -> None:
        for name in ("article", "supplier", "season", "source_sheet"):
            value = getattr(self, name)
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"Поле {name} должно быть непустой строкой.")
        if self.source_row < 1:
            raise ValueError("Номер строки Excel начинается с 1.")
        copied = dict(self.values)
        for key, value in copied.items():
            if not isinstance(key, str) or not key.strip():
                raise ValueError("Идентификатор поля должен быть непустой строкой.")
            if value is not None and not isinstance(
                value, str | int | float | bool | Decimal | date | datetime | ExcelErrorValue
            ):
                raise ValueError(f"Неподдерживаемый тип значения поля {key}.")
            if isinstance(value, float) and not math.isfinite(value):
                raise ValueError(f"Поле {key} содержит NaN или бесконечность.")
            if isinstance(value, Decimal) and not value.is_finite():
                raise ValueError(f"Поле {key} содержит недопустимое денежное значение.")
        object.__setattr__(self, "article", self.article.strip())
        object.__setattr__(self, "values", MappingProxyType(copied))
````

## src/kuchenland_importer/application/bind_photos.py

````python
"""Associate images in the declared photo column with article keys and build one asset per SKU."""

from dataclasses import dataclass, replace
from typing import Protocol

from kuchenland_importer.application.normalize_workbook import article_value
from kuchenland_importer.application.table_schema import ColumnRegistry, find_header
from kuchenland_importer.domain.photos import EmbeddedPhoto, PhotoExtraction
from kuchenland_importer.domain.product import PhotoAsset
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.workbook import SourceSheet, WorkbookPreview


class PhotoAssetStore(Protocol):
    def build(self, pictures: tuple[EmbeddedPhoto, ...]) -> PhotoAsset: ...


@dataclass(slots=True)
class BindPhotos:
    registry: ColumnRegistry
    store: PhotoAssetStore

    def execute(
        self, preview: WorkbookPreview, sheets: tuple[SourceSheet, ...], extracted: PhotoExtraction
    ) -> WorkbookPreview:
        cells: dict[tuple[str, int, int], str] = {}
        columns: dict[str, int] = {}
        issues = list(preview.issues)
        for sheet in sheets:
            try:
                header = find_header(sheet, self.registry)
                if header is None or not {"article", "photo"} <= set(header.columns):
                    continue
                column = header.columns.index("photo") + 1
                columns[sheet.name] = column
                for row_number, row in enumerate(sheet.rows, 1):
                    if row_number < header.row + header.depth:
                        continue
                    cell = row[header.columns.index("article")]
                    if cell.value is None or cell.value == "":
                        continue
                    article = article_value(cell)
                    base_key = (sheet.name, row_number, column)
                    if base_key in cells and cells[base_key] != article:
                        raise ValueError("Фото-ячейка объединяет разные артикулы.")
                    cells[base_key] = article
                    for r1, c1, r2, c2 in sheet.merges:
                        if r1 == row_number and c1 <= column <= c2:
                            for r in range(r1, r2 + 1):
                                for c in range(c1, c2 + 1):
                                    key = (sheet.name, r, c)
                                    if key in cells and cells[key] != article:
                                        raise ValueError("Фото-ячейка объединяет разные артикулы.")
                                    cells[key] = article
            except ValueError as error:
                issues.append(ImportIssue("PHOTO_TABLE", str(error), Severity.ERROR, sheet.name))
        grouped: dict[str, list[EmbeddedPhoto]] = {}
        for problem in extracted.problems:
            if problem.sheet in columns and (
                problem.column is None or (problem.sheet, problem.row, problem.column) in cells
                or problem.column == columns[problem.sheet]
            ):
                issues.append(ImportIssue(
                    "PHOTO_READ_FAILED", problem.message, Severity.ERROR,
                    f"{problem.sheet}:{problem.row or '?'}",
                ))
        for picture in extracted.pictures:
            key = (picture.sheet, picture.row, picture.column)
            target_article = cells.get(key)
            if target_article is None:
                continue  # Shapes outside the photo column are not product images.
            if picture.end_row is not None and any(
                cells.get((picture.sheet, r, picture.column)) not in {None, target_article}
                for r in range(picture.row + 1, picture.end_row + 1)
            ):
                issues.append(ImportIssue(
                    "PHOTO_SPANS_PRODUCTS", "Фото пересекает строки разных товаров.",
                    Severity.ERROR, f"{picture.sheet}:{picture.row}",
                ))
                continue
            grouped.setdefault(target_article, []).append(picture)
        products = []
        for product in preview.products:
            pictures = grouped.pop(product.article, [])
            asset = None
            if pictures:
                try:
                    asset = self.store.build(tuple(pictures))
                except Exception as error:
                    issues.append(ImportIssue(
                        "PHOTO_INVALID", str(error), Severity.ERROR,
                        f"{product.source_sheet}:{product.source_row}",
                    ))
            products.append(replace(product, photo=asset))
        for article in grouped:
            issues.append(ImportIssue(
                "PHOTO_WITHOUT_PRODUCT", "Изображение не связано с расчётной строкой.",
                Severity.WARNING, f"photo:{article}",
            ))
        return WorkbookPreview(tuple(products), tuple(issues))
````

## src/kuchenland_importer/infrastructure/photo_store.py

````python
"""Normalize visible pixels, build deterministic collages and persist content-addressed PNGs."""

import io
import math
import struct
import warnings
from hashlib import sha256
from pathlib import Path
from uuid import uuid4

from PIL import Image, ImageOps

from kuchenland_importer.domain.photos import EmbeddedPhoto
from kuchenland_importer.domain.product import PhotoAsset


def pixel_digest(image: Image.Image) -> str:
    rgba = image.convert("RGBA")
    return sha256(
        b"kuchenland-rgba-v1\0" + struct.pack("!II", *rgba.size) + rgba.tobytes()
    ).hexdigest()


def visible_image(photo: EmbeddedPhoto) -> Image.Image:
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        with Image.open(io.BytesIO(photo.data)) as source:
            if source.width * source.height > 40_000_000 or getattr(source, "n_frames", 1) > 1:
                raise ValueError("Фото превышает 40 Мп или содержит несколько кадров.")
            image = ImageOps.exif_transpose(source).convert("RGBA")
    left, top, right, bottom = photo.crop
    if min(photo.crop) < 0 or left + right >= 100000 or top + bottom >= 100000:
        raise ValueError("Обрезка фото не поддерживается или удаляет всё изображение.")
    if any(photo.crop):
        w, h = image.size
        box = (round(w * left / 100000), round(h * top / 100000),
               round(w * (1 - right / 100000)), round(h * (1 - bottom / 100000)))
        if box[0] >= box[2] or box[1] >= box[3]:
            raise ValueError("После обрезки у фото нет пикселей.")
        image = image.crop(box)
    if photo.flip_h:
        image = ImageOps.mirror(image)
    if photo.flip_v:
        image = ImageOps.flip(image)
    if not math.isfinite(photo.rotation):
        raise ValueError("Недопустимый угол поворота фото.")
    if photo.rotation % 360:
        image = image.rotate(-photo.rotation, resample=Image.Resampling.BICUBIC, expand=True)
    return image


class PhotoStore:
    def __init__(self, directory: Path) -> None:
        self.directory = directory.resolve()
        self.directory.mkdir(parents=True, exist_ok=True)

    def build(self, pictures: tuple[EmbeddedPhoto, ...]) -> PhotoAsset:
        if not pictures or len(pictures) > 16:
            raise ValueError("Для одного товара требуется от 1 до 16 изображений.")
        images = {}
        for picture in pictures:
            image = visible_image(picture)
            images[pixel_digest(image)] = image
        ordered = [images[key] for key in sorted(images)]
        if len(ordered) == 1:
            output = ordered[0]
        else:
            columns = math.ceil(math.sqrt(len(ordered)))
            rows = math.ceil(len(ordered) / columns)
            output = Image.new("RGBA", (columns * 512, rows * 512), "white")
            for index, image in enumerate(ordered):
                thumb = ImageOps.contain(image, (496, 496), Image.Resampling.LANCZOS)
                x = index % columns * 512 + (512 - thumb.width) // 2
                y = index // columns * 512 + (512 - thumb.height) // 2
                output.alpha_composite(thumb, (x, y))
        digest = pixel_digest(output)
        path = self.directory / f"{digest}.png"
        if path.exists():
            with Image.open(path) as cached:
                if pixel_digest(cached) != digest:
                    raise ValueError("Контрольная сумма сохранённого фото не совпадает.")
        else:
            partial = self.directory / f"{uuid4().hex}.part"
            try:
                with partial.open("xb") as stream:
                    output.save(stream, format="PNG")
                partial.rename(path)
            finally:
                partial.unlink(missing_ok=True)
        return PhotoAsset(path, digest, output.width, output.height, len(ordered))
````

## src/kuchenland_importer/infrastructure/excel/ooxml_package.py

````python
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
````

## src/kuchenland_importer/infrastructure/excel/drawing_photos.py

````python
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
````

## src/kuchenland_importer/infrastructure/excel/rich_photos.py

````python
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
````

## src/kuchenland_importer/infrastructure/excel/photo_reader.py

````python
"""Read image parts without opening or saving source books in Excel."""

from pathlib import Path
from uuid import uuid4

from kuchenland_importer.domain.photos import EmbeddedPhoto, PhotoExtraction, PhotoProblem
from kuchenland_importer.infrastructure.excel.drawing_photos import read_drawings
from kuchenland_importer.infrastructure.excel.ooxml_package import OOXMLPackage
from kuchenland_importer.infrastructure.excel.rich_photos import native_images, read_native


class PhotoReader:
    def read(self, path: Path, working_dir: Path | None = None) -> PhotoExtraction:
        if path.suffix.lower() in {".xls", ".xlsb"} and working_dir is not None:
            from kuchenland_importer.infrastructure.excel.legacy_photo_copy import (
                convert_photo_copy,
            )

            working_dir.mkdir(parents=True, exist_ok=True)
            copy = working_dir / f"photo-source-{uuid4().hex}.xlsx"
            try:
                convert_photo_copy(path, copy)
                return self.read(copy)
            finally:
                copy.unlink(missing_ok=True)
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Фото XLS/XLSB требуют конвертации рабочей копии через Excel.")
        package = OOXMLPackage(path)
        try:
            images = native_images(package)
            pictures: list[EmbeddedPhoto] = []
            problems: list[PhotoProblem] = []
            for name, part in package.sheets():
                drawing = read_drawings(package, name, part)
                pictures.extend(drawing.pictures)
                problems.extend(drawing.problems)
                pictures.extend(read_native(package, name, part, images))
            return PhotoExtraction(tuple(pictures), tuple(problems))
        finally:
            package.close()
````

## src/kuchenland_importer/infrastructure/excel/legacy_photo_copy.py

````python
"""Convert an old-format source to a disposable OOXML copy without saving the original."""

from pathlib import Path
from typing import Any


def convert_photo_copy(source: Path, destination: Path) -> None:
    import pythoncom
    import win32com.client

    if destination.exists():
        raise ValueError("Рабочая копия для фото уже существует.")
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app: Any = None
    book: Any = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        book = app.Workbooks.Open(
            str(source.resolve()), ReadOnly=True, UpdateLinks=0,
            Password="__unsupported_encryption__", AddToMru=False,
        )
        if any(book.Sheets(i).Type != -4167 for i in range(1, book.Sheets.Count + 1)):
            raise ValueError("Книга с листами макросов/диаграмм не поддерживается.")
        book.SaveAs(str(destination.resolve()), FileFormat=51)
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
````

## src/kuchenland_importer/infrastructure/excel/incell_writer.py

````python
"""Native picture conversion on an owned Excel working copy; never a floating fallback."""

from typing import Any

from PIL import Image

from kuchenland_importer.domain.product import PhotoAsset
from kuchenland_importer.infrastructure.photo_store import pixel_digest


class InCellPictureWriter:
    def place(self, sheet: Any, address: str, asset: PhotoAsset) -> None:
        with Image.open(asset.path) as image:
            if pixel_digest(image) != asset.pixel_sha256:
                raise ValueError("Фото изменилось после подготовки.")
        cell = sheet.Range(address)
        if cell.Count != 1 or cell.MergeCells:
            raise ValueError("Нативное фото требует одной необъединённой ячейки.")
        sheet.Activate()
        cell.Select()
        before = sheet.Shapes.Count
        picture = sheet.Shapes.AddPicture(
            str(asset.path), False, True, cell.Left + 0.1, cell.Top + 0.1,
            max(0.1, cell.Width - 0.2), max(0.1, cell.Height - 0.2),
        )
        try:
            picture.Select()
            picture.PlacePictureInCell()
            if sheet.Shapes.Count != before:
                raise ValueError("Excel не преобразовал фото в изображение в ячейке.")
            # Conversion initializes native images in a fresh book but resamples pixels.
            # Direct insertion then replaces that temporary image with the original PNG.
            cell.Select()
            cell.InsertPictureInCell(str(asset.path))
        except Exception:
            # The caller must discard the working copy if conversion failed.
            if sheet.Shapes.Count > before:
                picture.Delete()
            raise
````

## src/kuchenland_importer/infrastructure/preview_report.py

````python
"""Serialize typed values explicitly, preserving Excel error types."""

import json
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from kuchenland_importer.application.bind_photos import BindPhotos
from kuchenland_importer.application.normalize_workbook import NormalizeWorkbook
from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.application.route_workbook import RouteWorkbook
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.capture_loader import load_attachments
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader
from kuchenland_importer.infrastructure.photo_store import PhotoStore


def encode_value(value: object) -> object:
    if isinstance(value, ExcelErrorValue):
        return {"type": "excel_error", "code": value.code}
    if isinstance(value, datetime | date):
        return {"type": "date", "value": value.isoformat()}
    if isinstance(value, Decimal):
        return {"type": "decimal", "value": str(value)}
    raise TypeError(f"Нельзя сериализовать {type(value).__name__}")


@dataclass(frozen=True, slots=True)
class PreviewSummary:
    report: Path
    products: int
    errors: int
    routes_assigned: int = 0
    eligible_products: int = 0
    photos: int = 0


def create_preview(
    manifest: Path, columns: Path, output: Path, *, season_resolver: ResolveSeason | None = None,
    photo_directory: Path | None = None,
) -> PreviewSummary:
    try:
        attachments = load_attachments(manifest)
        registry = load_columns(columns)
        service = NormalizeWorkbook(registry)
        binder = BindPhotos(registry, PhotoStore(photo_directory)) if photo_directory else None
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ExcelInputError(f"Не удалось прочитать входные данные: {error}") from error
    result: dict[str, object] = {
        "schema_version": 1,
        "stage": "photo_preview" if binder is not None
        else "season_preview" if season_resolver is not None else "excel_preview",
    }
    files = []
    product_count = error_count = 0
    routes_assigned = eligible_products = skipped_attachments = 0
    photo_count = eligible_photos = 0
    for attachment in attachments:
        try:
            sheets = WorkbookReader().read(attachment.path)
            preview = service.execute(attachment, sheets)
            if binder is not None:
                preview = binder.execute(
                    preview, sheets, PhotoReader().read(attachment.path, photo_directory)
                )
            all_issues = preview.issues
            products = [
                {
                    "article": p.article,
                    "supplier": p.supplier,
                    "season": p.season,
                    "sheet": p.source_sheet,
                    "row": p.source_row,
                    "values": dict(p.values),
                }
                for p in preview.products
            ]
            file_photos = sum(p.photo is not None for p in preview.products)
            if binder is not None:
                for serialized, product in zip(products, preview.products, strict=True):
                    asset = product.photo
                    serialized["photo"] = (
                        {"path": str(asset.path), "pixel_sha256": asset.pixel_sha256,
                         "width": asset.width, "height": asset.height,
                         "image_count": asset.image_count}
                        if asset is not None else None
                    )
            if season_resolver is not None:
                routed = RouteWorkbook(season_resolver).execute(preview)
                all_issues = routed.issues
                for serialized, item in zip(products, routed.products, strict=True):
                    assignment = item.assignment
                    serialized["routing"] = (
                        {
                            "source_season": assignment.source_season,
                            "source_year": assignment.source_year,
                            "effective_year": assignment.effective_year,
                            "effective_season": assignment.effective_season,
                            "route_id": assignment.route.id,
                            "calculation_sheet": assignment.route.calculation_sheet,
                            "photo_sheet": assignment.route.photo_sheet,
                            "sales_start_month": assignment.sales_start_month,
                            "rule": assignment.rule,
                        }
                        if assignment is not None
                        else None
                    )
                routes_assigned += sum(p.assignment is not None for p in routed.products)
            issues = [
                {
                    "code": i.code,
                    "message": i.message,
                    "severity": i.severity.value,
                    "source": i.source,
                }
                for i in all_issues
            ]
            errors = sum(i.severity.value == "error" for i in all_issues)
            if errors:
                skipped_attachments += 1
            else:
                eligible_products += len(products)
                eligible_photos += file_photos
            files.append(
                {
                    "attachment": str(attachment.path),
                    "original_name": attachment.original_name,
                    "sha256": attachment.sha256,
                    "mail_subject": attachment.mail.subject,
                    "received_at": attachment.mail.received_at.isoformat(),
                    "products": products,
                    "issues": issues,
                    "eligible_for_import": errors == 0,
                }
            )
            product_count += len(products)
            photo_count += file_photos
            error_count += errors
        except Exception as error:
            skipped_attachments += 1
            files.append(
                {
                    "attachment": str(attachment.path),
                    "products": [],
                    "eligible_for_import": False,
                    "issues": [
                        {"code": "WORKBOOK_READ_FAILED", "message": str(error), "severity": "error"}
                    ],
                }
            )
            error_count += 1
    result["files"] = files
    result["product_count"] = product_count
    result["error_count"] = error_count
    if season_resolver is not None:
        result["routes_assigned"] = routes_assigned
        result["eligible_product_count"] = eligible_products
        result["skipped_attachment_count"] = skipped_attachments
    if binder is not None:
        result["photo_count"] = photo_count
        result["eligible_photo_count"] = eligible_photos
    partial = output.with_suffix(".json.part")
    try:
        with partial.open("x", encoding="utf-8") as stream:
            json.dump(
                result, stream, default=encode_value, ensure_ascii=False, indent=2, allow_nan=False
            )
        partial.rename(output)
    except (OSError, ValueError, TypeError) as error:
        raise ExcelInputError(f"Не удалось сохранить отчёт: {error}") from error
    return PreviewSummary(
        output, product_count, error_count, routes_assigned, eligible_products, photo_count
    )
````

## src/kuchenland_importer/presentation/cli.py

````python
"""Diagnostics and source collection; Excel import is not exposed yet."""

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from loguru import logger

from kuchenland_importer import __version__
from kuchenland_importer.app import Application
from kuchenland_importer.domain.errors import ApplicationError


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Kuchenland: диагностика и предпросмотр импорта")
    parser.add_argument(
        "command",
        nargs="?",
        default="diagnose",
        choices=(
            "diagnose", "capture-outlook", "preview-excel", "preview-seasons", "preview-photos"
        ),
    )
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument("--config", type=Path, required=True, help="Путь к настройкам TOML")
    parser.add_argument("--workbook", type=Path, help="Переопределить путь общей книги")
    parser.add_argument("--data-dir", type=Path, help="Переопределить рабочий каталог")
    parser.add_argument("--manifest", type=Path, help="Манифест полученных вложений")
    args = parser.parse_args(argv)
    # The executable owns its logger. Disable the default diagnostic stderr sink.
    logger.remove()
    app: Application | None = None
    try:
        if sys.version_info[:2] != (3, 12):
            print("Требуется Python 3.12.", file=sys.stderr)
            return 2
        app = Application.create(args.config, workbook=args.workbook, data_dir=args.data_dir)
        if args.command in {"preview-excel", "preview-seasons", "preview-photos"}:
            from uuid import uuid4

            from kuchenland_importer.application.resolve_season import ResolveSeason
            from kuchenland_importer.infrastructure.preview_report import create_preview

            if args.manifest is None:
                parser.error(f"{args.command} требует --manifest")
            resolver = (
                ResolveSeason(app.settings.routes) if args.command != "preview-excel" else None
            )
            prefix = "season-preview" if resolver is not None else "excel-preview"
            run_id = uuid4().hex
            photos = (
                app.paths.work / f"photo-preview-{run_id}"
                if args.command == "preview-photos" else None
            )
            if photos is not None:
                prefix = "photo-preview"
            preview = create_preview(
                args.manifest,
                args.config.resolve().parent / "columns.toml",
                app.paths.reports / f"{prefix}-{run_id}.json",
                season_resolver=resolver,
                photo_directory=photos,
            )
            print(f"Прочитано товаров: {preview.products}; ошибок: {preview.errors}")
            if photos is not None:
                print(f"Товаров с извлечённым фото: {preview.photos}")
            if resolver is not None:
                print(f"Назначено маршрутов: {preview.routes_assigned}")
                print(f"Товаров в корректных вложениях: {preview.eligible_products}")
                logger.info(
                    "Предпросмотр сезонов: товаров {}, маршрутов {}, ошибок {}",
                    preview.products, preview.routes_assigned, preview.errors,
                )
            print(f"Отчёт: {preview.report}")
            print("Это предварительный разбор; общая книга не изменялась.")
            return 1 if preview.errors else 0
        if args.command == "capture-outlook":
            outcome = app.capture_outlook()
            capture = outcome.capture
            print(f"Получено писем: {len(capture.mails)} из {capture.selected_count} элементов")
            print(f"Сохранено Excel-вложений: {len(capture.attachments)}")
            for issue in capture.issues:
                print(f"{issue.severity.value}: {issue.source}: {issue.message}")
            print(f"Отчёт: {outcome.manifest}")
            print("Получены исходные данные; общая книга не изменялась.")
            return 1 if capture.error_count or not capture.attachments else 0
        result = app.diagnose()
        print(f"Kuchenland Importer {__version__} — диагностика")
        for label, messages in (
            ("OK", result.checks),
            ("ПРЕДУПРЕЖДЕНИЕ", result.warnings),
            ("ОШИБКА", result.errors),
        ):
            for message in messages:
                print(f"{label}: {message}")
        return 1 if result.errors else 0
    except ApplicationError as error:
        if app is not None:
            logger.warning("Ошибка приложения: {}", error)
        print(f"Ошибка запуска: {error}", file=sys.stderr)
        return 2
    except Exception:
        if app is not None:
            logger.exception("Непредвиденная ошибка приложения")
            message = f"Непредвиденная ошибка. Журнал: {app.paths.logs / 'application.log'}"
        else:
            message = "Непредвиденная ошибка до открытия журнала. Проверьте настройки и доступы."
        print(message, file=sys.stderr)
        return 3
    finally:
        if app is not None:
            app.close()
````

## src/kuchenland_importer/__init__.py

````python
"""Kuchenland procurement importer."""

__version__ = "0.5.0"
````

## config/columns.toml

````toml
[fields]
article = ["Артикул", "Article", "SKU", "Item No", "Item Number"]
supplier = ["Поставщик", "Supplier", "Производитель", "Manufacturer"]
season = ["Сезон", "Season"]
name = ["Наименование", "Название", "Name", "Product Name", "Description of goods"]
sales_start = ["Старт продаж", "Начало продаж", "Sales Start"]
collection = ["Коллекция", "Collection"]
line = ["Линия", "Line"]
group = ["Группы", "группа", "Group"]
subgroup = ["Подгруппы", "подгруппа", "Subgroup"]
photo = ["Фото", "Photo", "Picture", "Image"]
description = ["Описание", "Description"]
model = ["Модель", "Model"]
size = ["Размер", "Size"]
manager = ["Ответственный менеджер", "Аналитик"]
quantity = ["Итого заказ", "Количество", "Quantity"]
retail_price = ["Предполагаемая цена", "Retail Price"]
purchase_price = ["Цена", "Цена, CNY", "Price", "Unit Price", "НОВЫЙ"]
currency = ["Валюта FOB", "Currency"]
moq = ["MOQ"]
target_price = ["TARGET PRICE"]
price_under_target = ["Цена под TARGET PRICE"]
comments = ["Комментарии", "Комментарий", "Comments"]
````

## pyproject.toml

````toml
[build-system]
requires = ["setuptools>=75,<83"]
build-backend = "setuptools.build_meta"

[project]
name = "kuchenland-importer"
version = "0.5.0"
description = "Windows procurement mail and Excel import tool"
requires-python = ">=3.12,<3.13"
dynamic = ["dependencies"]

[project.scripts]
kuchenland-importer = "kuchenland_importer.presentation.cli:main"

[tool.setuptools.packages.find]
where = ["src"]

[tool.setuptools.dynamic]
dependencies = {file = ["requirements.txt"]}

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
addopts = "-ra --strict-markers"

[tool.ruff]
target-version = "py312"
line-length = 100

[tool.ruff.lint]
select = ["E", "F", "I", "UP", "B"]

[tool.mypy]
python_version = "3.12"
mypy_path = "src"
files = ["src"]
strict = true

[[tool.mypy.overrides]]
module = ["pythoncom", "pywintypes", "win32com.*", "openpyxl.*"]
ignore_missing_imports = true
````

## tools/verify_incell_excel.py

````python
"""Acceptance check on a new disposable book; never opens the master workbook."""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from kuchenland_importer.domain.photos import EmbeddedPhoto  # noqa: E402
from kuchenland_importer.infrastructure.excel.incell_writer import InCellPictureWriter  # noqa: E402
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader  # noqa: E402
from kuchenland_importer.infrastructure.photo_store import PhotoStore  # noqa: E402


def main() -> None:
    import pythoncom
    import win32com.client
    from PIL import Image

    parser = argparse.ArgumentParser(description="Проверка нативных фото на новой тестовой книге")
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    root = args.directory.resolve()
    root.mkdir(parents=True, exist_ok=False)
    image = root / "input.png"
    Image.new("RGB", (90, 60), (240, 80, 40)).save(image)
    store = PhotoStore(root / "photos")
    asset = store.build((EmbeddedPhoto("Sheet1", 2, 2, image.read_bytes(), "QA"),))
    output = root / "native.xlsx"
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = book = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        book = app.Workbooks.Add()
        sheet = book.Worksheets(1)
        sheet.Name = "Photos"
        sheet.Range("A1:B1").Value = (("SKU", "Photo"),)
        sheet.Range("A2").Value = "QA-001"
        sheet.Range("B2").RowHeight = 80
        sheet.Range("B2").ColumnWidth = 18
        InCellPictureWriter().place(sheet, "B2", asset)
        book.SaveAs(str(output), FileFormat=51)
        book.Close(SaveChanges=False)
        book = app.Workbooks.Open(str(output), ReadOnly=True, UpdateLinks=0)
        assert book.Worksheets(1).Shapes.Count == 0
        pictures = PhotoReader().read(output).pictures
        assert len(pictures) == 1 and pictures[0].native
        assert (pictures[0].sheet, pictures[0].row, pictures[0].column) == ("Photos", 2, 2)
        assert store.build(pictures).pixel_sha256 == asset.pixel_sha256
        print(f"PASS: Excel {app.Version}, build {app.Build}; native image survives save/reopen.")
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
````

## tools/verify_legacy_photos.py

````python
"""Exercise XLS/XLSB value reading and photo conversion on new test files in installed Excel."""

import argparse
import sys
from hashlib import sha256
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader  # noqa: E402
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader  # noqa: E402


def main() -> None:
    import pythoncom
    import win32com.client
    from PIL import Image

    parser = argparse.ArgumentParser(
        description="Проверка старых форматов на новых тестовых книгах"
    )
    parser.add_argument("directory", type=Path)
    root = parser.parse_args().directory.resolve()
    root.mkdir(parents=True, exist_ok=False)
    image = root / "image.png"
    Image.new("RGB", (40, 30), "blue").save(image)
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = book = None
    paths = []
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        for suffix, file_format in (("xls", 56), ("xlsb", 50)):
            book = app.Workbooks.Add()
            sheet = book.Worksheets(1)
            sheet.Range("A1:C1").Value = (("SKU", "Photo", "Name"),)
            sheet.Range("A2").Value = "QA-001"
            sheet.Range("C2").Value = "QA product"
            cell = sheet.Range("B2")
            sheet.Shapes.AddPicture(str(image), False, True, cell.Left, cell.Top, 30, 20)
            path = root / f"source.{suffix}"
            book.SaveAs(str(path), FileFormat=file_format)
            book.Close(SaveChanges=False)
            book = None
            paths.append(path)
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
    for path in paths:
        before = sha256(path.read_bytes()).digest()
        values = WorkbookReader().read(path)
        photos = PhotoReader().read(path, root / "copies")
        assert values[0].rows[1][0].value == "QA-001"
        assert len(photos.pictures) == 1 and not photos.problems
        assert (photos.pictures[0].row, photos.pictures[0].column) == (2, 2)
        assert sha256(path.read_bytes()).digest() == before
        assert not list((root / "copies").glob("photo-source-*.xlsx"))
        print(f"PASS {path.suffix}: values and photo read; source SHA-256 unchanged.")


if __name__ == "__main__":
    main()
````

## tests/unit/test_photos.py

````python
import io
from dataclasses import replace
from datetime import UTC, datetime
from pathlib import Path

import pytest
from PIL import Image, UnidentifiedImageError

from kuchenland_importer.application.bind_photos import BindPhotos
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.photos import (
    EmbeddedPhoto,
    PhotoAction,
    PhotoExtraction,
    photo_action,
)
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.workbook import SourceCell, SourceSheet, WorkbookPreview
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.photo_store import PhotoStore, pixel_digest, visible_image

CONFIG = Path(__file__).resolve().parents[2] / "config" / "columns.toml"


def payload(color: str = "red", compression: int = 6) -> bytes:
    stream = io.BytesIO()
    Image.new("RGB", (40, 20), color).save(stream, format="PNG", compress_level=compression)
    return stream.getvalue()


def picture(color: str = "red", **kwargs):
    return EmbeddedPhoto("Images", 2, 2, payload(color), "source", **kwargs)


def source_product(path: Path, article: str = "A") -> ProductRecord:
    mail = MailMetadata("entry", "store", "", datetime(2026, 10, 9, tzinfo=UTC), "", "")
    attachment = SavedAttachment(mail, 1, "source.xlsx", path.resolve(), "0" * 64)
    return ProductRecord(article, "Vendor", "Весна", attachment, "Items", 2, {"custom": "keep"})


def photo_sheet() -> SourceSheet:
    return SourceSheet("Images", tuple(tuple(SourceCell(v) for v in r) for r in (
        ("SKU", "Photo", "Name"), ("A", None, "First"), ("B", None, "Second")
    )))


def test_hash_ignores_encoding_and_collage_deduplicates(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    first = store.build((picture(),))
    second = store.build((replace(picture(), data=payload(compression=0)),))
    duplicate = store.build((picture(), picture()))
    assert first == second == duplicate
    assert not list(tmp_path.glob("*.part"))


def test_collage_is_independent_of_drawing_order(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    result = store.build((picture("red"), picture("blue")))
    reverse = store.build((picture("blue"), picture("red")))
    assert result == reverse and (result.width, result.height) == (1024, 512)
    assert result.image_count == 2
    with Image.open(result.path) as image:
        assert len(image.getcolors(maxcolors=10000)) >= 3  # Both images and white margins.


def test_visible_crop_and_rotation_change_pixels() -> None:
    original = picture()
    cropped = visible_image(replace(original, crop=(50000, 0, 0, 0)))
    rotated = visible_image(replace(original, rotation=90))
    assert cropped.size == (20, 20) and rotated.size == (20, 40)
    assert pixel_digest(cropped) != pixel_digest(visible_image(original))


@pytest.mark.parametrize("crop", [(100000, 0, 0, 0), (-1, 0, 0, 0), (50000, 0, 50000, 0)])
def test_invalid_crop_is_not_silently_ignored(crop: tuple[int, int, int, int]) -> None:
    with pytest.raises(ValueError, match="Обрезка"):
        visible_image(picture(crop=crop))


def test_corrupt_payload_and_cache_are_detected(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    with pytest.raises(UnidentifiedImageError):
        store.build((replace(picture(), data=b"broken"),))
    asset = store.build((picture(),))
    Image.new("RGB", (40, 20), "blue").save(asset.path)
    with pytest.raises(ValueError, match="сумма"):
        store.build((picture(),))


@pytest.mark.parametrize("count", [0, 17])
def test_photo_count_limit(tmp_path: Path, count: int) -> None:
    with pytest.raises(ValueError, match="от 1 до 16"):
        PhotoStore(tmp_path).build((picture(),) * count)


def test_photo_update_policy(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    old = store.build((picture("red"),))
    new = store.build((picture("blue"),))
    assert photo_action(None, None) == PhotoAction.NONE
    assert photo_action(None, new) == PhotoAction.ADD
    assert photo_action(old, old) == PhotoAction.KEEP
    assert photo_action(old, None) == PhotoAction.KEEP
    assert photo_action(old, new) == PhotoAction.REPLACE


def test_binding_uses_article_and_photo_column(tmp_path: Path) -> None:
    first = source_product(tmp_path / "book.xlsx")
    second = source_product(tmp_path / "book.xlsx", "B")
    # Source products and picture records have unrelated orders.
    pictures = (replace(picture("blue"), row=3), picture(), replace(picture("green"), column=3))
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((second, first), ()), (photo_sheet(),), PhotoExtraction(pictures)
    )
    assert not result.issues
    assert all(p.photo is not None for p in result.products)
    assert result.products[0].photo.pixel_sha256 != result.products[1].photo.pixel_sha256
    assert dict(result.products[0].values) == {"custom": "keep"}


def test_photo_spanning_products_blocks_attachment(tmp_path: Path) -> None:
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((source_product(tmp_path / "book.xlsx"),), ()),
        (photo_sheet(),), PhotoExtraction((picture(end_row=3),)),
    )
    assert result.issues[0].code == "PHOTO_SPANS_PRODUCTS"
    assert result.products[0].photo is None


def test_missing_photo_is_a_valid_absence(tmp_path: Path) -> None:
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((source_product(tmp_path / "book.xlsx"),), ()),
        (photo_sheet(),), PhotoExtraction(()),
    )
    assert not result.issues and result.products[0].photo is None


def test_corrupt_photo_becomes_explicit_error(tmp_path: Path) -> None:
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((source_product(tmp_path / "book.xlsx"),), ()),
        (photo_sheet(),), PhotoExtraction((replace(picture(), data=b"broken"),)),
    )
    assert result.issues[0].code == "PHOTO_INVALID"
    assert result.products[0].photo is None
````

## tests/unit/test_incell_writer.py

````python
from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from kuchenland_importer.domain.photos import EmbeddedPhoto
from kuchenland_importer.infrastructure.excel.incell_writer import InCellPictureWriter
from kuchenland_importer.infrastructure.photo_store import PhotoStore


def asset_at(root: Path):
    file = root / "image.png"
    Image.new("RGB", (40, 30), "red").save(file)
    return PhotoStore(root / "photos").build((
        EmbeddedPhoto("Images", 2, 2, file.read_bytes(), "QA"),
    ))


def test_changed_photo_is_rejected_before_excel_access(tmp_path: Path) -> None:
    asset = asset_at(tmp_path)
    Image.new("RGB", (40, 30), "blue").save(asset.path)
    with pytest.raises(ValueError, match="изменилось"):
        InCellPictureWriter().place(None, "B2", asset)


def test_merged_destination_is_rejected(tmp_path: Path) -> None:
    sheet = SimpleNamespace(Range=lambda _: SimpleNamespace(Count=1, MergeCells=True))
    with pytest.raises(ValueError, match="необъединённой"):
        InCellPictureWriter().place(sheet, "B2", asset_at(tmp_path))


def test_noop_conversion_raises_and_removes_temporary_floating_picture(tmp_path: Path) -> None:
    asset = asset_at(tmp_path)
    shapes = SimpleNamespace(Count=0)
    deleted = []

    def add(*args):
        shapes.Count += 1
        return SimpleNamespace(
            Select=lambda: None, PlacePictureInCell=lambda: None,
            Delete=lambda: deleted.append(True),
        )

    shapes.AddPicture = add
    cell = SimpleNamespace(Count=1, MergeCells=False, Left=0, Top=0, Width=30, Height=20,
                           Select=lambda: None)
    sheet = SimpleNamespace(Range=lambda _: cell, Shapes=shapes, Activate=lambda: None)
    with pytest.raises(ValueError, match="не преобразовал"):
        InCellPictureWriter().place(sheet, "B2", asset)
    assert deleted == [True]
````

## tests/integration/test_photo_preview.py

````python
import io
import json
from datetime import UTC, datetime
from hashlib import sha256
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image

from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader
from kuchenland_importer.presentation.cli import main

CONFIG = Path(__file__).resolve().parents[2] / "config" / "default.toml"


def test_cli_extracts_collage_and_keeps_inputs_unchanged(tmp_path: Path) -> None:
    book = Workbook()
    book.active.append(["SKU", "Supplier", "Season"])
    book.active.append(["A", "Vendor", "Весна 2027"])
    sheet = book.create_sheet("Images")
    sheet.append(["SKU", "Photo", "Name"])
    sheet.append(["A", None, "First"])
    for color in ("red", "blue"):
        buffer = io.BytesIO()
        Image.new("RGB", (40, 20), color).save(buffer, format="PNG")
        sheet.add_image(ExcelImage(buffer), "B2")
    source = tmp_path / "input.xlsx"
    book.save(source)
    before = source.read_bytes()
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps({
        "schema_version": 1, "stage": "mail_capture", "mails": [{
            "entry_id": "entry", "store_id": "store", "subject": "Subject",
            "sender_name": "", "sender_address": "",
            "received_at": datetime(2026, 10, 9, tzinfo=UTC).isoformat(),
        }], "attachments": [{"mail_index": 0, "attachment_index": 1,
            "original_name": source.name, "path": source.name,
            "sha256": sha256(before).hexdigest()}],
    }), encoding="utf-8")
    runtime = tmp_path / "runtime"
    assert main(["preview-photos", "--config", str(CONFIG), "--manifest", str(manifest),
                 "--data-dir", str(runtime)]) == 0
    data = json.loads(next((runtime / "reports").glob("photo-preview-*.json")).read_text('utf-8'))
    assert data["stage"] == "photo_preview"
    assert data["photo_count"] == 1 and data["eligible_photo_count"] == 1
    product = data["files"][0]["products"][0]
    asset = product["photo"]
    assert (asset["width"], asset["height"]) == (1024, 512)
    assert asset["image_count"] == 2
    assert Path(asset["path"]).is_file()
    assert product["routing"]["route_id"] == "spring"
    assert source.read_bytes() == before


def test_one_cell_anchor_detects_visible_row_span(tmp_path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.append(["SKU", "Photo", "Name"])
    sheet.append(["A", None, "First"])
    sheet.append(["B", None, "Second"])
    buffer = io.BytesIO()
    Image.new("RGB", (40, 35), "red").save(buffer, format="PNG")
    sheet.add_image(ExcelImage(buffer), "B2")
    path = tmp_path / "spanning.xlsx"
    book.save(path)
    result = PhotoReader().read(path)
    assert result.pictures[0].row == 2 and result.pictures[0].end_row == 3
````

## tests/integration/test_native_photos.py

````python
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
````

## README.md

````markdown
# Kuchenland Importer

Windows-приложение для импорта товаров из выделенных писем классического Outlook
в локальную общую книгу Excel. Python 3.12, Office Professional Plus 2024 / Microsoft 365.

## Текущий статус

Реализованы этапы 1–5: основа, получение писем, разбор Excel, маршрутизация сезонов
и извлечение/подготовка фотографий.
На реальных XLSX-вложениях из 13 MSG прочитано 90 товаров. Автоматические тесты проверены;
приёмка на настоящем Outlook отложена до устройства с настроенной учётной записью.
Запись товаров и фото в общую книгу, GUI и EXE ещё не реализованы.
Приложение пока не готово к производственному использованию.

## Установка для разработки

В PowerShell из корня проекта:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

`requirements.txt` содержит рабочие зависимости; `requirements-dev.txt` добавляет
проверки качества. Если присутствует `requirements-lock.txt`, он фиксирует версии
проверенного окружения и устанавливается вместо `requirements-dev.txt` для его
точного воспроизведения. `tkinter` входит в стандартную Windows-установку Python.

## Запуск этапа 1

```powershell
.\.venv\Scripts\python.exe main.py --config config/default.toml
```

Команда проверяет настройки, создаёт рабочие каталоги и журнал, проверяет наличие
библиотек. Если общая книга не выбрана, сообщает предупреждение. Не подключается
к Outlook/Excel, не изменяет и не сохраняет книги.

Для проверки своей книги передайте абсолютный путь:

```powershell
.\.venv\Scripts\python.exe main.py --config config/default.toml --workbook "D:\вайбкодин\Почта_сводник\=РАССЧИТАНО= ВЕСНА  ВЕСНА-ЛЕТО и ПАСХА 08.10.2026.xlsx"
```

Для персональных настроек скопируйте `config/default.toml` в `config/local.toml`
и задайте `application.workbook`. `local.toml` исключён из Git.
Каждый конфигурационный файл — полный документ; автоматического слияния нет.
Относительные пути в TOML считаются от папки файла настроек. Относительные пути
аргументов CLI считаются от текущего рабочего каталога.

`--data-dir` переопределяет рабочий каталог. По умолчанию при запуске из исходников
это `.runtime/`, содержащий `logs`, `backup`, `work`, `reports`.
Для будущей установленной версии пользовательский каталог будет определён на этапе 8.

Коды завершения: `0` — проверка успешна, возможны предупреждения;
`1` — диагностика обнаружила проблемы; `2` — ошибка запуска/настроек/аргументов;
`3` — непредвиденная ошибка. `--version` работает без конфигурации.

## Получение писем — этап 2

На компьютере с настроенным классическим Outlook откройте основное окно и выделите
письма в списке. Затем запустите из корня проекта:

```powershell
.\.venv\Scripts\python.exe main.py capture-outlook --config config/default.toml
```

Outlook автоматически не запускается. Для этой команды общая книга не требуется.
Читаются идентификаторы, тема, дата получения и отправитель. Excel-вложения
`.xlsx`, `.xlsm`, `.xls`, `.xlsb` сохраняются в `.runtime/work/<run_id>/mail-NNNN/`.
Формат файла пока определяется по расширению; его содержимое проверяется на этапе 3.
Небезопасные исходные имена никогда не используются для построения пути сохранения.

`manifest.json` содержит данные писем, позиции в исходном выделении, оригинальные
имена вложений, относительные пути, SHA-256 и ошибки. Время получения записывается
в UTC. Адрес Exchange может быть служебным адресом, который вернул Outlook;
он не подменяет поставщика из таблицы. Тело письма не читается и не сохраняется.
Отчёт содержит служебные данные и хранится локально вне Git.

Корректные вложения сохраняются даже при ошибках других вложений. Пустые и
неполностью сохранённые файлы не включаются в результат; временные файлы удаляются,
а невозможность их удаления явно сообщается как ошибка. Повторный запуск получает
свой каталог; предыдущие результаты не перезаписываются.
Это получение исходных данных, а не импорт товаров в общую книгу.

Для `capture-outlook`: `0` — есть сохранённые вложения и нет ошибок;
`1` — частичные ошибки либо нет Excel-вложений; `2` — запуск/Outlook недоступен;
`3` — непредвиденная ошибка. Предупреждения также выводятся при успешном запуске.
Если невозможно сохранить JSON-отчёт, полученные файлы остаются в каталоге запуска.
Предоставленные MSG пока не открываются этой командой: она читает выделение Outlook.

## Предварительный разбор Excel — этап 3

Передайте манифест, созданный командой capture-outlook:

```powershell
.\.venv\Scripts\python.exe main.py preview-excel --config config/default.toml --manifest "ПОЛНЫЙ_ПУТЬ_К_manifest.json"
```

Словарь столбцов находится в `config/columns.toml` рядом с TOML-настройками приложения.
Артикул/Article/SKU/Item No приводятся к одному полю. Поддерживаются многоуровневые
и объединённые заголовки исследованных шаблонов. Порядок именованных столбцов
не влияет на поиск полей. Нераспознанные поля сохраняются с префиксом source:;
безымянные столбцы получают исходный номер, для их автоматического переноса потребуется
явное сопоставление на этапе 6. Неоднозначные обязательные столбцы не угадываются.

Поставщик и сезон берутся из строки расчётной таблицы. Поля фото-листа соединяются
по артикулу и сохраняются с префиксом photo:. Для извлечения изображений используется
отдельная команда preview-photos (этап 5).
Числовые артикулы с простым форматом 00000 восстанавливают ведущие нули;
текстовые артикулы сохраняют регистр и нули. Распознаваемые цены становятся Decimal,
даты начала продаж — date; пояснительные текстовые значения сохраняются.

Ошибки Excel остаются типизированными `{type: excel_error, code: ...}`, а не обычным
текстом. Они отражаются предупреждениями и не отменяют строку по согласованному правилу.
Если отсутствует сохранённый результат формулы, строка отмечается ошибкой:
openpyxl не вычисляет формулы. Пустой текстовый результат формулы — допустимое значение.

XLSX/XLSM читаются без сохранения книги. XLS/XLSB используют отдельный экземпляр
Excel, отключают VBA через AutomationSecurity и открывают книгу ReadOnly без обновления
ссылок. Нативная совместимость XLS/XLSB, политики Excel 4.0 macros и особенности
защищённых файлов требуют проверки на целевой машине; листы макросов не поддерживаются.
При отсутствии подходящего Excel файл попадёт в отчёт с ошибкой.

SHA-256 вложений проверяется перед чтением. JSON-отчёт записывается в .runtime/reports;
повреждённая книга не блокирует другие вложения. Счётчик товаров — число прочитанных
записей предпросмотра, не число импортированных товаров. eligible_for_import отражает
валидность чтения; финальная проверка сезона, дублей между письмами и записи ещё впереди.
Исходные файлы и общая книга не изменяются. Лимиты: 50 МБ на OOXML-файл,
100000 строк и 512 столбцов на лист; произвольные шаблоны требуют настройки профиля.

Для QA предоставленных Unicode MSG без Outlook есть `tools/extract_msg_samples.py`.
Это ограниченная утилита исследования материалов, а не универсальный MSG-импортёр.
Она сохраняет вложения и совместимый манифест в новый каталог, читая Windows OLE
только для чтения. Исходные MSG проверяются SHA-256. Команда для предоставленной папки:

```powershell
.\.venv\Scripts\python.exe tools/extract_msg_samples.py "D:\вайбкодин\Почта_сводник" ".runtime\msg-qa-new"
```

## Предпросмотр сезонных маршрутов — этап 4

Команда `preview-seasons` читает исходные вложения из манифеста и для каждой строки
назначает расчётную вкладку и соответствующий лист ФОТО. Для уже извлечённых
предоставленных материалов запуск из корня проекта:

```powershell
.\.venv\Scripts\python.exe main.py preview-seasons --config config/default.toml --manifest .runtime/msg-stage3-verified/manifest.json
```

Для писем Outlook передайте путь к `manifest.json` своего запуска capture-outlook.
Настроенные псевдонимы сопоставляются без учёта регистра, с нормализацией пробелов
и вариантов тире. Поддержан исследованный формат `(У) 2. ВЕСНА-ЛЕТО 2027`:
пометки `(У)`/`(Д)`, порядковый номер и один четырёхзначный год отделяются от названия.
Другие пометки, диапазоны годов и составные сезоны без псевдонима не угадываются.
Точное совпадение с настроенным псевдонимом имеет приоритет перед разбором пометок.

Маршрут с id `winter` при месяце начала продаж 12 меняется на маршрут `spring`.
Эти идентификаторы используются декабрьским правилом; если настроен winter,
нужен spring. Имена вкладок и псевдонимы обоих маршрутов можно менять.
Месяц определяется из date/datetime, календарных дат ДД.ММ.ГГГГ, ДД/ММ/ГГГГ,
ГГГГ-ММ-ДД, месяца с годом ММ.ГГГГ/ГГГГ-ММ, ДД.ММ и полного русского/английского
названия месяца, в том числе с годом. Нераспознанный текст, числа без формата даты,
пустые значения и ошибки Excel не интерпретируются произвольно.
При неопределённом начале продаж Зима остаётся на зимней паре с предупреждением.

НГ и ВСЕСЕЗОННЫЙ сейчас не настроены по решению пользователя. Любой неизвестный
сезон делает всё вложение непригодным для будущего импорта, включая остальные
корректные строки этого вложения. Другие вложения продолжают обрабатываться.
Предпросмотр сохраняет все прочитанные строки и объясняет ошибки.

Отчёт `.runtime/reports/season-preview-<run_id>.json` содержит `routing` каждой
строки: исходный сезон/год, итоговый сезон/год, месяц, правило и точную пару вкладок.
Неопределённый маршрут — null. `routes_assigned` считает назначения, в том числе
в ошибочных вложениях; `eligible_product_count` считает только товары из полностью
корректных вложений. Эти числа не означают запись товаров в общую книгу.
Коды завершения: 0 — нет ошибок, 1 — есть ошибочные вложения/строки,
2 — ошибка входных данных или настроек, 3 — непредвиденный сбой.

Исходная строка и её поля сохраняются без изменения; итоговый сезон находится
в назначении маршрута и должен использоваться будущим адаптером записи.
Исходный год сохраняется отдельно. По согласованному правилу декабрьский переход
повышает его на один: Зима 2026 → Весна 2027. Если в названии года нет, используется
год распознанной даты начала продаж. Если нет обоих, Весна назначается без года
с предупреждением TARGET_SEASON_YEAR_UNKNOWN. Для остальных маршрутов год названия
сохраняется; при его отсутствии год из даты автоматически не добавляется.
Наличие/схема вкладок общей книги пока не проверяются, новые листы не создаются;
это этап 6. Изображения и выбор самого нового письма относятся к этапам 5–6.

## Настройки сезонов

Каждый `[[routes]]` содержит уникальный `id`, список названий `aliases`,
имя расчётной вкладки `calculation_sheet` и листа `photo_sheet`.
Можно добавлять новые маршруты без изменения моделей и исходного кода.
Неоднозначные названия, повторные назначения вкладок и опечатки в ключах отклоняются.

Настроены Весна/Весна-лето, Пасха, Лето/Лето-осень, Осень, Зима.
В исходной общей книге зимней пары нет: на этом этапе конфигурация только описывает
её, а вкладки не создаются. Создание пары по выбранному шаблону и редактор сезонов
относятся к этапам 6–7.

Адаптеры чтения `.xlsx`, `.xlsm`, `.xls`, `.xlsb` реализованы на этапе 3.
Реальные вложения проверены для XLSX; тестовые XLS/XLSB проверены через установленный
Microsoft 365. Приёмка произвольных файлов и Office 2024 остаётся отдельной проверкой.
Файлы с паролем не обещаются к поддержке без отдельного решения.

## Фотографии — этап 5

Для предоставленного набора материалов:

```powershell
.\.venv\Scripts\python.exe main.py preview-photos --config config/default.toml --manifest .runtime/msg-stage3-verified/manifest.json
```

Для Outlook используйте манифест своего запуска capture-outlook. Команда объединяет
проверки таблиц, сезонов и фото. Вложения и общая книга не изменяются. JSON-отчёт
находится в `.runtime/reports/photo-preview-<run_id>.json`, подготовленные PNG —
в `.runtime/work/photo-preview-<run_id>/`.

Изображения привязываются к артикулу строки по именованному столбцу Фото/Photo/Image.
Картинки в других столбцах не считаются фотографиями товара. Поддержаны DrawingML
oneCellAnchor/twoCellAnchor и нативные localImage в ячейке с rich-data metadata.
Используются сохранённые пиксели с учётом EXIF, обрезки, поворота и отражений.
Фото, пересекающее строки разных артикулов, — ошибка, а не случайный выбор товара.
Внешние ссылки не загружаются из сети. Группы/абсолютные привязки требуют профиля
и дают явную ошибку на фото-листе; повреждённые изображения блокируют своё вложение.

Несколько разных фото собираются в коллаж по решению пользователя. Одинаковые
нормализованные изображения удаляются из набора. Порядок частей коллажа определяется
хешем, поэтому изменение порядка объектов не создаёт ложную замену. Сетка использует
плитки 512×512, белый фон и сохранение пропорций; для одиночного изображения размер
не уменьшается. `image_count` — число разных изображений в подготовленном фото.
Лимиты: до 16 исходных изображений товара, 20 МБ на картинку, 40 Мп на декодирование.
Объём распакованного OOXML ограничен 256 МБ/10000 частей, XML-часть — 16 МБ.

Хеш включает размеры и RGBA-пиксели, поэтому перекодирование без изменения пикселей
не меняет результат. Решение фото: add/replace/keep/none. Если нового фото нет,
прежнее сохраняется; одинаковое не заменяется; изменённое заменяется. Ошибка чтения
не считается отсутствием нового фото: такое вложение исключается из импорта.
Сравнение с фото общей книги и применение решения будут подключены на этапе 6.

XLS/XLSB для чтения фото конвертируются в отдельную временную XLSX-копию через Excel.
Источник открывается ReadOnly без обновления ссылок, VBA отключён; копия удаляется
после чтения. Источник не сохраняется. Политики XLM/защищённых файлов остаются
ограничениями целевой среды. Сеансы чтения/записи Excel должны выполняться последовательно.

Адаптер InCellPictureWriter работает только на рабочей копии и одной необъединённой
ячейке. Сначала конвертация инициализирует нативный формат, затем прямая вставка
сохраняет исходные пиксели PNG. Если Excel не поддерживает операции, возникает ошибка;
подмена плавающим изображением не применяется. При ошибке рабочую копию нужно отбросить.
Сохранённый результат должен проверяться перед заменой общей книги на этапе 6.

Нативные проверки на новых тестовых книгах, в ранее не существовавших каталогах:

```powershell
.\.venv\Scripts\python.exe tools/verify_incell_excel.py .tmp/incell-acceptance-new
.\.venv\Scripts\python.exe tools/verify_legacy_photos.py .tmp/legacy-acceptance-new
```

В текущем Microsoft 365 16.0.20430.20092 подтверждены правильная ячейка, нативная
метаинформация, отсутствие плавающей картинки и точный хеш после save/reopen.
Проверки на Office Professional Plus 2024 ещё требуются. Макросы VBA не используются.

## Проверки

```powershell
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe -m ruff check .
.\.venv\Scripts\python.exe -m mypy
```

Тесты проверяют конфигурацию и модели, ошибки запуска, журналы, COM-жизненный цикл,
смешанное выделение, дубли, безопасные имена, частичные ошибки вложений и отчёт,
поиск заголовков, нормализацию значений и сохранение ошибок Excel.
COM-объекты в автоматических тестах заменены управляемыми тестовыми объектами.
Настоящая интеграция с Outlook ещё не испытана; порядок приёмки описан отдельно.

## Документы

- `docs/specification.md` — согласованные бизнес-правила и открытые вопросы.
- `docs/architecture/decisions.md` — границы слоёв и решения по Office/сохранению.
- `docs/roadmap.md` — этапы и критерии приёмки.
- `docs/stage-1-code.md` — полные тексты кода и настроек этапа 1 для ревью.
- `docs/stage-1-validation.md` — результаты и границы выполненных проверок.
- `docs/stage-2-code.md` — полные актуальные файлы этапа 2 для ревью.
- `docs/stage-2-validation.md` — проверки и приёмка на другом компьютере.
- `docs/stage-3-code.md` — полные новые и изменённые файлы этапа 3.
- `docs/stage-3-validation.md` — результаты чтения и ограничения проверки.
- `docs/stage-4-code.md` — полные новые и изменённые файлы маршрутизации.
- `docs/stage-4-validation.md` — проверки сезонных правил на тестах и вложениях.
- `docs/stage-5-code.md` — полные новые и изменённые файлы обработки фотографий.
- `docs/stage-5-validation.md` — автоматические и нативные проверки этапа 5.

Рабочие письма и книги, логи, резервные копии, `.venv` и личные настройки исключены
из Git. После каждого завершённого этапа и успешных проверок код публикуется
в [MailExcel2Git](https://github.com/MaxGarAI/MailExcel2Git) в ветку
`codex/development`. Порядок публикации закреплён в `AGENTS.md`.
````

## docs/specification.md

````markdown
# Спецификация импорта

Версия 0.5. Уточнения пользователя от 09.10.2026 (Europe/Moscow).

## Согласованные правила

1. Целевая среда: Windows, Python 3.12, Office Professional Plus 2024 и Microsoft 365.
   Для COM-сценария требуется классический Outlook; точные сборки испытываются отдельно.
2. Общая книга локальная. Каждый товар представлен на расчётном сезонном листе
   и соответствующем листе ФОТО. Лист `удалено` не участвует в поиске/изменениях.
3. Предварительно артикул уникален во всей книге как идентификатор товара.
   Его две строки в сезонной паре не считаются двумя разными товарами.
4. Источник поставщика — таблица во вложении. Отправитель письма не подменяет поставщика.
   Дата — дата получения выбранного письма Outlook; тема берётся из этого письма.
5. Сезон определяется на уровне строки товара. Одно вложение может иметь разные сезоны.
   Весна/Весна-лето → Весна, Пасха → Пасха, Лето/Лето-осень → Лето,
   Осень → Осень, Зима → Зима. Зима с началом продаж в декабре → Весна.
6. Сезоны и пары вкладок должны расширяться настройками и впоследствии через GUI.
   Создание новой пары потребует определения шаблона; неизвестный сезон нельзя
   безусловно направлять в произвольную вкладку.
7. Входящие данные вставляются как значения, не как формулы поставщика.
   При обновлении заменяются импортируемые поля; формулы в столбцах, которых нет
   во вложении, сохраняются. Служебные ячейки вне строк товаров сохраняются.
   Физический способ замены должен обеспечивать это правило и сохранность ссылок.
8. Изменившаяся фотография заменяет старую; одинаковая не заменяется;
   при отсутствии новой фотографии прежняя сохраняется. Требуется настоящее
   изображение в ячейке Excel, без молчаливой подмены плавающей картинкой.
9. При нескольких письмах с одним артикулом выигрывает самое новое по ReceivedTime.
   Ошибка в новом вложении не означает разрешение незаметно применить старое письмо.
10. Корректные вложения импортируются, ошибочные пропускаются. Частичный результат
    явно отражается в отчёте. Системный сбой записи требует восстановления книги.
11. При смене сезона товар и фото переносятся на новую пару, старые записи удаляются.
12. `Цена под TARGET PRICE` сохраняется и импортируется при наличии.
13. Целевые форматы вложений: `.xlsx`, `.xlsm`, `.xls`, `.xlsb`.
    Для старых/бинарных форматов требуется Excel-адаптер; макросы не запускаются.
14. По уточнению от 09.10.2026 строки с ошибками Excel импортируются с сохранением
    ошибки в ячейке. Ошибка Excel — отдельный тип значения, не текст и не ноль.
    Отсутствие кеша формулы не считается результатом: требуется расчёт в Excel.
15. НГ и ВСЕСЕЗОННЫЙ пока не маршрутизируются. Неизвестный сезон — ошибка,
    всё содержащее его вложение пропускается, остальные вложения обрабатываются.
16. Зима без определимого начала продаж (пустое значение, ошибка Excel,
    неоднозначный текст) остаётся на зимней паре с предупреждением.
17. При декабрьском переходе год повышается: Зима 2026 → Весна 2027.
    Исходный сезон/год сохраняются отдельно для проверки происхождения данных.
18. Несколько разных изображений в фото-ячейке товара собираются в один коллаж.
    Идентичные изображения не дублируются; картинки вне столбца Фото не выбираются
    автоматически как дополнительные фотографии товара.

## Факты предварительного анализа

В папке D:\вайбкодин\Почта_сводник обнаружено 13 MSG и общая книга.
В исследованной книге 9 листов, 275 файлов изображений. Расчётные и фото-листы
содержат формулы (в частности VLOOKUP на фото-листах). Пользователь уточнил,
что формулы в столбцах вне импортируемого вложения сохраняются.
Существующие ошибки формул не исправляются как побочный эффект импорта.
Схемы фото-листов не одинаковы; некоторые заголовки повторяются (Коллекция).
Нужны профили и контекст заголовка, а не слепой словарь название → колонка.

Предварительный анализ не заменяет полного разбора вложений и проверки Excel.

## Вопросы для соответствующих этапов

- Если даты писем равны и значения расходятся: требуется явный конфликт,
  пока не утверждено правило приоритета.
- Регистр и допустимая нормализация артикулов; значимость пробелов/дефисов;
  дубли внутри вложения и совпадение артикула у разных поставщиков.
- Составные сезоны вне утверждённых псевдонимов.
- Название поля начала продаж и поставщика в реальных вложениях.
- Что делать с ручными значениями в столбцах, отсутствующих во вложении:
  пользователь явно согласовал сохранение формул, но не всех ручных значений.
- Шаблон новой расчётной/фото-пары; какие формулы и оформление переносить.
- Число пользователей и работа с открытой/несохранённой книгой. До согласования
  запись в открытую пользователем книгу не включается.
- Профили группированных изображений и защищённых файлов. На этапе 5 неоднозначная
  привязка/повреждение фото отклоняет вложение; отсутствие фото допускается.
````

## docs/architecture/decisions.md

````markdown
# Архитектурные решения

## Слои и зависимости

Domain зависит только от стандартной библиотеки: dataclass-модели, инварианты,
происхождение записи, отчёт. Application содержит сценарии и порты Protocol.
Infrastructure реализует работу с конфигурацией, Office, файлами и журналированием.
Presentation преобразует пользовательский ввод в вызовы приложения.
`app.py` — composition root; его диагностический сценарий относится к этапу 1.

Каждый адаптер отвечает за одну внешнюю систему. Outlook/Excel не импортируются
в Domain и не передают COM-объекты в модели. Объекты Office живут в потоке с
инициализированным COM; GUI будет получать события через очередь и tkinter.after.

## Формат настроек

TOML читается штатным tomllib Python 3.12. Загрузка строгая: неизвестные ключи,
неверные типы, конфликты псевдонимов и вкладок приводят к ConfigurationError.
Настройки и маршруты неизменяемы. Сезон — расширяемая строка, не закрытый enum.
Выбор пары реализован отдельным сервисом этапа 4; загрузчик только проверяет настройки.

## Модели и данные

Артикул — строка, ведущие нули и регистр сохраняются. Текстовые поля не преобразуются
в формулы Excel. ProductRecord содержит копию словаря значений с запретом изменений,
ссылку на письмо/вложение и координаты исходной строки.
Фото имеет хеш нормализованных пикселей; извлечение реализовано на этапе 5.
ReceivedTime должен преобразовываться адаптером в datetime с часовым поясом.
Денежные значения могут храниться в Decimal; NaN и бесконечность недопустимы.

## Office и запись

openpyxl предназначен для чтения подходящих OOXML-вложений. Общую книгу с объектами
и нативными изображениями в ячейке будет сохранять Excel через xlwings/pywin32.
Конкретный API изображений и сохранение после повторного открытия проверяются
на Office 2024 и Microsoft 365 до приёмки этапа 5–6.
xlwings pictures.add с anchor само по себе не доказывает режим изображения в ячейке.

Источник формул поставщика импортируется как вычисленные значения; если кеша нет,
на этапе 3 строка получает ошибку. Источник нужно пересчитать в Excel и получить заново.
Политика защищённых файлов будет определена по фактическим входам.

Две строки сезонной пары представляют один товар. Индексы дублей учитывают эту пару,
а не объявляют зеркальные строки конфликтом. Изменение сезона — согласованный перенос
обоих представлений с сохранением прежнего фото, если новое не пришло.

Корректные вложения отделяются от ошибочных до записи. План, резервная копия,
проверка состояния исходной книги, запись в рабочую копию и проверка сохранённого
результата должны предшествовать замене общего файла. Гарантии восстановления
подтверждаются fault-injection испытаниями; Excel сам по себе не транзакционная БД.
В этапе 1 книги только проверяются на чтение и никогда не изменяются.

## Логи и доставка

Loguru: UTF-8, ротация по размеру, хранение по сроку. diagnose/backtrace выключены,
чтобы исключения не выводили значения локальных переменных. Содержимое писем не
логируется автоматически. CLI отключает стандартный stderr-sink Loguru; каждый
экземпляр приложения при завершении освобождает только свой файловый sink.

Настоящие письма, книги и резервные копии не включаются в Git. Фиксированные версии
проверенного окружения хранятся отдельно от диапазонов совместимости.
Сборка EXE, пользовательские каталоги, обновления и подпись относятся к этапу 8.

## Этап 2: источник Outlook

MailSource и CaptureManifestStore — порты Application. CaptureMail создаёт отдельный
каталог UUID, вызывает источник и сохраняет манифест. Данные MailCapture содержат
идентичности, позицию в выделении, пути и контрольные суммы; COM-объектов в них нет.

Адаптер подключается через GetActiveObject только к уже запущенному классическому
Outlook. COM инициализируется STA в вызывающем потоке и освобождается в finally.
Состав выделения считывается до обработки вложений. Объекты другого класса
пропускаются с предупреждением. Ошибки отдельных писем/вложений становятся issues.

Дата получения читается из PR_MESSAGE_DELIVERY_TIME (PT_SYSTIME) и сохраняется
как UTC; локальный ReceivedTime не используется для повторной конвертации времени.
Имена сохранённых файлов генерируются по индексам; исходное имя — только метаданные.
Временный файл публикуется после проверки размера и вычисления SHA-256.
Сохранение вложений не означает, что Excel-содержимое уже проверено или импортировано.

JSON-отчёт связывает сохранённые файлы с письмами. Сбой отчёта оставляет вложения
для восстановления и вызывает явную ошибку. Тема, адрес и идентификаторы писем
сохраняются только в локальном манифесте; логи содержат номера позиций и коды ошибок.
Тела писем не читаются. Новый запуск не перезаписывает предыдущий.

Из-за отсутствия настроенного Outlook здесь интеграционная приёмка перенесена
на другое устройство по согласованию с пользователем. Это не блокирует разработку
следующих этапов, но остаётся обязательной частью приёмки приложения.

## Этап 3: чтение Excel и нормализация

WorkbookReader возвращает модели SourceSheet/SourceCell без объектов библиотек.
OOXML читается дважды: формулы и сохранённые значения. Это позволяет различать
пустой результат строковой формулы и отсутствующий кеш. LegacyReader изолирует
COM-чтение XLS/XLSB, открывает отдельный экземпляр Excel и закрывает его в finally.

NormalizeWorkbook и ColumnRegistry работают только с моделями Domain. Словарь
псевдонимов загружается отдельным TOML-адаптером. Заголовки связываются с полями,
неизвестные столбцы сохраняются; фото-описания соединяются по артикулу.
Нестандартные шаблоны требуют явного профиля, а не угадывания обязательных полей.

ExcelErrorValue хранит код ошибки как самостоятельный тип. Он допустим в строке
товара по согласованному правилу; будущий адаптер записи должен восстановить ошибку
Excel, а не записать её строковое название. Decimal и даты также сериализуются
с явным типом. Предпросмотр сохраняет источник, строку, тему и дату письма.

Ошибки чтения отдельной книги не останавливают другие книги. Повреждение манифеста
или нарушение контрольных сумм останавливает предпросмотр до чтения: происхождение
данных должно быть достоверным. Предпросмотр не выбирает сезонную пару и не записывает
общую книгу. Метка eligible_for_import означает только отсутствие ошибок чтения;
окончательную готовность определят следующие этапы.

## Этап 4: маршрутизация строк

SeasonRoute перенесён в Domain; Infrastructure сохраняет совместимый импорт этого
типа для прежнего кода. Application не зависит от настроек Infrastructure.
ResolveSeason строит неизменяемый индекс настроенных псевдонимов, разбирает наблюдавшиеся
пометки/номер/год и возвращает SeasonAssignment с причиной выбора пары.
sales_month отвечает только за распознавание однозначного месяца, без обращения к Office.

RouteWorkbook сохраняет исходные ProductRecord и ошибки всех строк. Неизвестное
назначение представлено явно; одна такая строка исключает всё вложение из кандидатов.
Зимняя строка без определимого месяца допускается с предупреждением согласно решению
пользователя. Модели сезонных назначений не меняют исходные данные и не создают вкладки.

preview-seasons использует тот же адаптер чтения, что preview-excel. Отчёт содержит
исходный сезон и назначение отдельно, чтобы будущий writer использовал итоговый сезон
и точные имена пары, сохраняя происхождение данных. Подсчёт назначений отделён от числа
товаров в пригодных вложениях. Ни один из этих счётчиков не означает сохранение книги.
Год назначения при переходе из декабря повышается на один по уточнению пользователя.
Сначала берётся год из сезона, при его отсутствии — из даты начала продаж.
Если ни один источник года недоступен, назначение Весны сохраняется с предупреждением.

## Этап 5: фотографии

Domain описывает исходное изображение с координатами/обрезкой и решение add/replace/keep/none.
BindPhotos зависит от порта PhotoAssetStore, связывает фото-ячейку с артикулом и
не получает объекты Pillow/Excel. Повреждение фото становится ошибкой вложения;
отсутствие изображения остаётся допустимым None и сохраняет старое фото при обновлении.

OOXMLPackage читает части и внутренние связи с лимитами размера. DrawingPhotos
разбирает обычные привязки, RichPhotos — цепочку метаинформации нативных ячеек.
PhotoReader объединяет результаты; legacy-книги преобразуются только в временные
копии. Источники не сохраняются, внешние изображения не скачиваются.

PhotoStore реализует Pillow-преобразования и контентное хранение PNG. Размеры и RGBA
пиксели входят в хеш. Несколько разных фотографий образуют детерминированный коллаж;
исходное расположение вне фото-столбца не используется для угадывания привязки.

InCellPictureWriter — COM-адаптер для рабочей копии будущего writer. В проверенном
Microsoft 365 необходимы активация/выделение и начальная конвертация; затем прямая
вставка восстанавливает исходные пиксели, поскольку одна конвертация растеризует
картинку до размеров shape. Сохранённая rich-data связь и хеш проверены после reopen.
Это проверенное поведение конкретной сборки, не гарантия всей матрицы Office.
Все сеансы Excel в сценарии должны выполняться последовательно; общую книгу
можно заменять только после полной проверки рабочей копии на этапе 6.
````

## docs/roadmap.md

````markdown
# План и критерии приёмки

На 09.10.2026 реализованы этапы 1–5. Их проверки и ограничения зафиксированы
в docs/stage-N-validation.md. Нативные изображения и тестовые legacy-книги проверены
в текущем Microsoft 365; Outlook и Office 2024 остаются в целевой приёмке.
Следующий этап — запись в общую книгу с резервным копированием и восстановлением.

1. Основа: конфигурация, модели, logger, CLI-диагностика. Проверки настроек,
   моделей и ошибок запуска; ruff, mypy. Не открывает COM и не изменяет книги.
2. Outlook: выбранные письма классического Outlook, ReceivedTime, безопасное
   сохранение вложений, EntryID/StoreID. Проверка пустого/смешанного выделения,
   кириллицы, одинаковых имён вложений; разбор предоставленных MSG на копиях.
3. Вложения Excel: реальные профили поставщиков, синонимы и контекст заголовков,
   значения формул, типы и неизвестные поля. Проверка переставленных столбцов,
   двух «Коллекция», ведущих нулей и поддерживаемых форматов.
4. Сезон: строковые сезоны, год, декабрьское правило, конфликты, настраиваемые
   пары. Тесты всех согласованных маршрутов; неизвестное значение явно сообщается.
5. Фото: связь с артикулом, хеш пикселей, изменённое/отсутствующее фото,
   нативное изображение в ячейке. Проверка в реальном Excel, после save/reopen.
6. Общая книга: частичный импорт корректных вложений, выбор нового письма,
   две строки товара, перенос сезона, сохранение формул вне импортируемых полей,
   backup и восстановление. Проверка повторного импорта, конфликтов, сбоя записи,
   сортировки и фильтрации фотографий; испытания на копиях пользовательской книги.
7. GUI: импорт, настройки сезонных пар и шаблонов, прогресс, результаты,
   запрет параллельных запусков и определённая отмена. Проверка отзывчивости.
8. EXE: воспроизводимая сборка и версия, каталоги пользователя, инструкция,
   проверка Windows без Python и матрица Office 2024 / Microsoft 365.

К следующему этапу переходить после проверки текущего. Завершение этапа 1 не
означает готовность импорта. Готовность коммерческой версии подтверждается
приёмкой реальных сценариев и восстановлением после сбоев на целевых машинах.
````

## docs/stage-5-validation.md

````markdown
# Этап 5: проверка фотографий

09.10.2026 (Europe/Moscow), Python 3.12.8, Windows, приложение 0.5.0.

- Pytest: 144 passed; Ruff: All checks passed; Mypy strict: 49 source files.
- Проверены хеш независимо от PNG-сжатия, удаление одинаковых изображений,
  детерминированный коллаж, обрезка/поворот, повреждение изображения и кеша,
  ограничения числа картинок, отсутствие фото и политика add/replace/keep/none.
- Проверены привязка по артикулу независимо от порядка товаров, игнорирование
  соседних столбцов, пересечение строк разных товаров и отчёт CLI.
  Для oneCellAnchor конечная строка вычисляется по высотам строк и размерам рисунка.
- Для native images проверены цепочка vm → valueMetadata → futureMetadata →
  rich value → структура → relationship → media, порядок ключей и типов metadata,
  отрицательные/выходящие индексы и отказ от внешних изображений.
- Проверены отказ при изменённом PNG, объединённой ячейке и тихой неудаче конвертации
  Excel: временная плавающая картинка не выдаётся за нативное изображение.

## Предоставленные материалы

На 13 реальных XLSX-вложениях прочитаны 90 товаров и подготовлены фото для всех 90.
7 подготовленных фото — коллажи из разных изображений одного товара. Новых ошибок
фотографий нет; остаются 3 ранее согласованные ошибки неизвестных сезонов в 2 вложениях.
Из полностью корректных вложений получены 82 товара с 82 фото для будущего импорта.
Счётчики не означают запись в общую книгу и не учитывают ещё выбор нового письма.
Все 32 типизированных значения ошибок Excel сохранились.
SHA-256 совпал до/после для 27 файлов: 13 MSG, 13 вложений и общей книги.

Локальный проверенный отчёт:
`.runtime/reports/photo-preview-441b2796fcfb4c5a88bdbe60b1f51414.json`.
Исходные и извлечённые изображения/отчёты не включаются в GitHub.

## Настоящий Excel

Установлен O365ProPlusRetail 16.0.20430.20092 (Microsoft 365), Excel Version 16.0,
Build 20430. Проверка tools/verify_incell_excel.py прошла на новой тестовой книге
`.tmp/incell-stage5-roundtrip/native.xlsx` после сохранения и повторного открытия:
картинка в Photos!B2, rich-data metadata, Shapes.Count = 0, хеш пикселей совпал.

Обнаружено и учтено поведение этой сборки: без выделения конвертация может быть
тихой неудачей; обычная конвертация shape изменяет размеры пикселей; прямая вставка
в новую книгу до инициализации native images может завершиться COM-ошибкой.
Проверенный порядок — выделение/конвертация, затем прямая вставка исходного PNG.
Адаптер не принимает растеризованную плавающую картинку как равноценный результат.

tools/verify_legacy_photos.py прошёл для новых `.xls` и `.xlsb` в
`.tmp/legacy-stage5-roundtrip`: чтение значений и фото, правильная привязка 2/2,
неизменный SHA-256 источников, удаление временных XLSX-копий. Сеансы выполняются
последовательно; произвольные реальные legacy-шаблоны ещё требуют приёмки.

## Границы этапа

Общая книга не изменена. Нативный адаптер проверен на тестовой книге; применение
политики фото к общей книге, удаление прежних floating shapes, перенос между сезонами,
сортировка/фильтры товарных строк и транзакционное сохранение относятся к этапу 6.
Office Professional Plus 2024 и другие сборки Microsoft 365 здесь не испытаны.
Группы, absoluteAnchor, внешние изображения, отрицательная обрезка и анимированные
изображения не заявлены как поддержанные; ошибочные вложения явно отклоняются.
Зашифрованные книги и политики XLM требуют целевой проверки.

[Описание настоящего режима Place in Cell от Microsoft](https://support.microsoft.com/en-gb/excel/insert-picture-in-cell-in-excel).
````
