# Этап 3: полные файлы для ревью

Версия 0.3.0. Снимок новых и изменённых файлов. Рабочие файлы проекта — источник истины; письма, книги и персональные отчёты сюда не включены.

## src/kuchenland_importer/__init__.py

````python
"""Kuchenland procurement importer."""

__version__ = "0.3.0"
````

## src/kuchenland_importer/domain/cell_values.py

````python
"""Typed Excel errors must not be converted to strings or missing values."""

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal


@dataclass(frozen=True, slots=True)
class ExcelErrorValue:
    code: str

    def __post_init__(self) -> None:
        if not self.code.startswith("#"):
            raise ValueError("Недопустимый код ошибки Excel.")


type CellValue = str | int | float | bool | Decimal | date | datetime | ExcelErrorValue | None
````

## src/kuchenland_importer/domain/workbook.py

````python
"""Read-only workbook snapshots and normalized preview results."""

from dataclasses import dataclass

from kuchenland_importer.domain.cell_values import CellValue
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.report import ImportIssue


@dataclass(frozen=True, slots=True)
class SourceCell:
    value: CellValue
    number_format: str = "General"
    missing_formula_cache: bool = False


@dataclass(frozen=True, slots=True)
class SourceSheet:
    name: str
    rows: tuple[tuple[SourceCell, ...], ...]
    merges: tuple[tuple[int, int, int, int], ...] = ()


@dataclass(frozen=True, slots=True)
class WorkbookPreview:
    products: tuple[ProductRecord, ...]
    issues: tuple[ImportIssue, ...]
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

    def __post_init__(self) -> None:
        if not self.path.is_absolute() or self.width < 1 or self.height < 1:
            raise ValueError("Фотография должна иметь абсолютный путь и положительный размер.")
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

## src/kuchenland_importer/domain/errors.py

````python
"""Errors that can be translated into actionable user messages."""


class ApplicationError(Exception):
    """Base class for expected application failures."""


class ConfigurationError(ApplicationError):
    """Configuration is missing, inconsistent, or malformed."""


class StartupError(ApplicationError):
    """Application resources cannot be initialized."""


class OutlookError(ApplicationError):
    """Classic Outlook is unavailable or its selection cannot be obtained."""


class AttachmentError(ApplicationError):
    """An attachment could not be safely saved and verified."""


class CaptureStorageError(ApplicationError):
    """Source capture directories or manifest could not be written."""


class ExcelInputError(ApplicationError):
    """Excel preview inputs or report storage are invalid or unavailable."""
````

## src/kuchenland_importer/application/table_schema.py

````python
"""Configurable aliases; ambiguous duplicate headers cannot silently overwrite values."""

import re
from dataclasses import dataclass

from kuchenland_importer.domain.workbook import SourceSheet


def header_key(value: str) -> str:
    return re.sub(r"[\s.]+", " ", value.casefold().replace("ё", "е")).strip()


@dataclass(frozen=True, slots=True)
class ColumnRegistry:
    aliases: dict[str, str]

    def resolve(self, value: str) -> str | None:
        return self.aliases.get(header_key(value))


@dataclass(frozen=True, slots=True)
class TableHeader:
    row: int
    depth: int
    columns: tuple[str, ...]


def find_header(sheet: SourceSheet, registry: ColumnRegistry) -> TableHeader | None:
    for index, row in enumerate(sheet.rows[:50]):
        fields = [
            registry.resolve(str(cell.value)) if cell.value is not None else None for cell in row
        ]
        if "article" not in fields or sum(field is not None for field in fields) < 3:
            continue
        depth = 1
        if index + 1 < len(sheet.rows):
            next_row = sheet.rows[index + 1]
            sku = fields.index("article")
            vertical_header = any(
                r1 <= index + 1 and r2 >= index + 2 for r1, _, r2, _ in sheet.merges
            )
            label_count = sum(
                isinstance(cell.value, str) and cell.value.strip() != "" for cell in next_row
            )
            known_secondary = any(
                registry.resolve(str(cell.value)) is not None
                for cell in next_row
                if cell.value is not None
            )
            if next_row[sku].value is None and (
                vertical_header or label_count >= 3 or label_count >= 2 and known_secondary
            ):
                depth = 2
        names: list[str] = []
        counts: dict[str, int] = {}
        for col in range(len(row)):
            labels = []
            for r in range(index + 1, index + depth + 1):
                value = sheet.rows[r - 1][col].value
                for r1, c1, r2, c2 in sheet.merges:
                    if r1 <= r <= r2 and c1 <= col + 1 <= c2:
                        value = sheet.rows[r1 - 1][c1 - 1].value
                        break
                if value is not None and str(value).strip() and str(value).strip() not in labels:
                    labels.append(str(value).strip())
            field = registry.resolve(labels[-1]) if labels else None
            if field is None and len(labels) == 1:
                field = registry.resolve(labels[0])
            key = field or "source:" + (
                " / ".join(header_key(label) for label in labels) or f"unnamed-column-{col + 1}"
            )
            if key in names and not key.startswith("source:"):
                # Photo templates distinguish a descriptive Collection from Line/Collection/Name.
                if (
                    key == "collection"
                    and "photo" in fields
                    and col > 0
                    and col + 1 < len(fields)
                    and fields[col - 1] == "line"
                    and fields[col + 1] == "name"
                ):
                    previous = names.index(key)
                    names[previous] = "source:descriptive-collection"
                    counts[key] = 0
                elif key not in {"article", "supplier", "season", "name", "collection"}:
                    previous = names.index(key)
                    names[previous] = "source:" + key + "#1"
                    key = "source:" + key + "#2"
                else:
                    raise ValueError(f"Неоднозначный столбец {key} на листе {sheet.name}.")
            counts[key] = counts.get(key, 0) + 1
            names.append(key if counts[key] == 1 else f"{key}#{counts[key]}")
        return TableHeader(index + 1, depth, tuple(names))
    return None
````

## src/kuchenland_importer/application/value_normalization.py

````python
"""Normalize identifiable numeric/date representations without inventing missing values."""

import re
from datetime import datetime
from decimal import Decimal

from kuchenland_importer.domain.cell_values import CellValue

PRICE_FIELDS = {"purchase_price", "retail_price", "target_price", "price_under_target"}


def normalize_value(field: str, value: CellValue) -> CellValue:
    if field in PRICE_FIELDS and not isinstance(value, bool):
        if isinstance(value, int | float | Decimal):
            return Decimal(str(value))
        if isinstance(value, str):
            text = value.strip().replace("\u00a0", "").replace(" ", "")
            if re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", text):
                return Decimal(text.replace(",", "."))
    if field == "sales_start" and isinstance(value, str):
        for pattern in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value.strip(), pattern).date()
            except ValueError:
                continue
    return value
````

## src/kuchenland_importer/application/normalize_workbook.py

````python
"""Normalize by field names; join photo descriptions by article, not row position."""

import re
from dataclasses import dataclass

from kuchenland_importer.application.table_schema import (
    ColumnRegistry,
    TableHeader,
    find_header,
)
from kuchenland_importer.application.value_normalization import normalize_value
from kuchenland_importer.domain.cell_values import CellValue, ExcelErrorValue
from kuchenland_importer.domain.mail import SavedAttachment
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.workbook import SourceCell, SourceSheet, WorkbookPreview


def article_value(cell: SourceCell) -> str:
    value = cell.value
    if isinstance(value, str):
        if not value.strip():
            raise ValueError("Пустой артикул.")
        return value.strip()
    if type(value) is int or isinstance(value, float) and value.is_integer():
        pattern = cell.number_format
        if re.fullmatch(r"0+", pattern):
            return str(int(value)).zfill(len(pattern))
        return str(int(value))
    raise ValueError("Артикул должен быть текстом или целым числом.")


def row_values(row: tuple[SourceCell, ...], header: TableHeader) -> dict[str, CellValue]:
    if any(cell.missing_formula_cache for cell in row):
        raise ValueError("У формулы нет сохранённого результата; требуется расчёт в Excel.")
    return {
        name: normalize_value(name, cell.value)
        for name, cell in zip(header.columns, row, strict=True)
    }


@dataclass(slots=True)
class NormalizeWorkbook:
    registry: ColumnRegistry

    def execute(
        self, attachment: SavedAttachment, sheets: tuple[SourceSheet, ...]
    ) -> WorkbookPreview:
        issues: list[ImportIssue] = []
        primary: list[tuple[SourceSheet, TableHeader]] = []
        photos: dict[str, dict[str, CellValue]] = {}
        for sheet in sheets:
            try:
                header = find_header(sheet, self.registry)
                if header is None:
                    continue
                if {"article", "supplier", "season"} <= set(header.columns):
                    primary.append((sheet, header))
                elif {"article", "photo"} <= set(header.columns):
                    for row in sheet.rows[header.row + header.depth - 1 :]:
                        sku = row[header.columns.index("article")]
                        if sku.value is None or sku.value == "":
                            continue
                        article = article_value(sku)
                        if article in photos:
                            raise ValueError(f"Повтор артикула на фото-листах: {article}")
                        photos[article] = row_values(row, header)
            except ValueError as error:
                issues.append(ImportIssue("SHEET_SCHEMA", str(error), Severity.ERROR, sheet.name))
        products = []
        seen: set[str] = set()
        for sheet, header in primary:
            for row_number, row in enumerate(
                sheet.rows[header.row + header.depth - 1 :], header.row + header.depth
            ):
                sku = row[header.columns.index("article")]
                if sku.value is None or sku.value == "":
                    continue
                source = f"{sheet.name}:{row_number}"
                try:
                    article = article_value(sku)
                    if article in seen:
                        raise ValueError(f"Повтор артикула в расчётных таблицах: {article}")
                    values = row_values(row, header)
                    supplier, season = values["supplier"], values["season"]
                    if not isinstance(supplier, str) or not supplier.strip():
                        raise ValueError("Не определён поставщик из строки таблицы.")
                    if not isinstance(season, str) or not season.strip():
                        raise ValueError("Не определён сезон из строки таблицы.")
                    values.update(
                        {"photo:" + key: value for key, value in photos.get(article, {}).items()}
                    )
                    products.append(
                        ProductRecord(
                            article,
                            supplier.strip(),
                            season.strip(),
                            attachment,
                            sheet.name,
                            row_number,
                            values,
                        )
                    )
                    seen.add(article)
                    if any(isinstance(value, ExcelErrorValue) for value in values.values()):
                        issues.append(
                            ImportIssue(
                                "EXCEL_ERROR_VALUE",
                                "Ошибки Excel сохранены как ошибки.",
                                Severity.WARNING,
                                source,
                            )
                        )
                except ValueError as error:
                    issues.append(ImportIssue("ROW_INVALID", str(error), Severity.ERROR, source))
        if not primary:
            issues.append(
                ImportIssue(
                    "NO_PRODUCT_TABLE",
                    "Нет таблицы с артикулом, поставщиком и сезоном.",
                    Severity.ERROR,
                    attachment.original_name,
                )
            )
        for article in photos.keys() - seen:
            issues.append(
                ImportIssue(
                    "PHOTO_WITHOUT_PRODUCT",
                    "Фото-строка не связана с товаром.",
                    Severity.WARNING,
                    f"photo:{article}",
                )
            )
        return WorkbookPreview(tuple(products), tuple(issues))
````

## src/kuchenland_importer/infrastructure/excel/__init__.py

````python
"""Read-only workbook adapters. No destination workbook writes at this stage."""
````

## src/kuchenland_importer/infrastructure/excel/reader.py

````python
"""Read OOXML values and formula-cache status without saving source workbooks."""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import cast

from openpyxl import load_workbook

from kuchenland_importer.domain.cell_values import CellValue, ExcelErrorValue
from kuchenland_importer.domain.workbook import SourceCell, SourceSheet


def cell_value(value: object, is_error: bool = False) -> CellValue:
    if is_error:
        return ExcelErrorValue(str(value))
    if value is None or isinstance(value, str | int | float | bool | Decimal | date | datetime):
        return cast(CellValue, value)
    raise ValueError(f"Неподдерживаемый тип ячейки: {type(value).__name__}")


class WorkbookReader:
    def read(self, path: Path) -> tuple[SourceSheet, ...]:
        if path.suffix.lower() in {".xls", ".xlsb"}:
            from kuchenland_importer.infrastructure.excel.legacy_reader import read_legacy

            return read_legacy(path)
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Неизвестный формат Excel.")
        if path.stat().st_size > 50 * 1024 * 1024:
            raise ValueError("Вложение превышает лимит чтения 50 МБ.")
        formulas = load_workbook(path, data_only=False, keep_links=False)
        try:
            values = load_workbook(path, data_only=True, keep_links=False)
            try:
                sheets = []
                for sheet in formulas:
                    if sheet.max_row > 100_000 or sheet.max_column > 512:
                        raise ValueError("Лист превышает лимит 100000 строк / 512 столбцов.")
                    cached = values[sheet.title]
                    rows = tuple(
                        tuple(
                            SourceCell(
                                cell_value(
                                    ""
                                    if cell.data_type == "f"
                                    and cached.cell(cell.row, cell.column).value is None
                                    and cached.cell(cell.row, cell.column).data_type in {"str", "s"}
                                    else cached.cell(cell.row, cell.column).value,
                                    cached.cell(cell.row, cell.column).data_type == "e",
                                ),
                                cell.number_format,
                                cell.data_type == "f"
                                and cached.cell(cell.row, cell.column).value is None
                                and cached.cell(cell.row, cell.column).data_type
                                not in {"str", "s"},
                            )
                            for cell in row
                        )
                        for row in sheet.iter_rows()
                    )
                    merges = tuple(
                        (r.min_row, r.min_col, r.max_row, r.max_col)
                        for r in sheet.merged_cells.ranges
                    )
                    sheets.append(SourceSheet(sheet.title, rows, merges))
                return tuple(sheets)
            finally:
                values.close()
        finally:
            formulas.close()
````

## src/kuchenland_importer/infrastructure/excel/legacy_reader.py

````python
"""Isolated Excel process for legacy/binary sources; never attach to user Excel."""

from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from kuchenland_importer.domain.cell_values import CellValue, ExcelErrorValue
from kuchenland_importer.domain.workbook import SourceCell, SourceSheet

ERROR_CODES = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
}


def read_legacy(path: Path) -> tuple[SourceSheet, ...]:
    from kuchenland_importer.infrastructure.excel.reader import cell_value

    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app: Any = None
    book: Any = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        app.AskToUpdateLinks = False
        book = app.Workbooks.Open(
            str(path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            Password="__unsupported_encryption__",
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        if any(sheet.Type != -4167 for sheet in book.Sheets):
            raise ValueError("Листы макросов и иные типы листов не поддерживаются.")
        result = []
        for sheet in book.Worksheets:
            used = sheet.UsedRange
            max_row = used.Row + used.Rows.Count - 1
            max_col = used.Column + used.Columns.Count - 1
            if max_row > 100_000 or max_col > 512:
                raise ValueError("Лист превышает лимит размера.")
            rows = []
            merges = set()
            for row in range(1, max_row + 1):
                cells = []
                for col in range(1, max_col + 1):
                    cell = sheet.Cells(row, col)
                    value = cell.Value
                    if (
                        isinstance(value, int)
                        and value - (-2146828288) in ERROR_CODES
                        and app.WorksheetFunction.IsError(cell)
                    ):
                        normalized: CellValue = ExcelErrorValue(ERROR_CODES[value - (-2146828288)])
                    else:
                        normalized = cell_value(value)
                    cells.append(SourceCell(normalized, str(cell.NumberFormat)))
                    if cell.MergeCells:
                        area = cell.MergeArea
                        merges.add(
                            (
                                area.Row,
                                area.Column,
                                area.Row + area.Rows.Count - 1,
                                area.Column + area.Columns.Count - 1,
                            )
                        )
                rows.append(tuple(cells))
            result.append(SourceSheet(str(sheet.Name), tuple(rows), tuple(sorted(merges))))
        return tuple(result)
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                book = None
                app = None
                pythoncom.CoUninitialize()
````

## src/kuchenland_importer/infrastructure/excel/columns.py

````python
"""Load the configurable field dictionary without embedding IO in business rules."""

import tomllib
from pathlib import Path

from kuchenland_importer.application.table_schema import ColumnRegistry, header_key


def load_columns(path: Path) -> ColumnRegistry:
    with path.open("rb") as stream:
        data = tomllib.load(stream)
    if set(data) != {"fields"} or not isinstance(data["fields"], dict):
        raise ValueError("Ожидается таблица fields в настройках столбцов.")
    aliases = {}
    for field, names in data["fields"].items():
        if not isinstance(names, list) or not names:
            raise ValueError("Поле должно иметь список псевдонимов.")
        for name in names:
            if not isinstance(name, str) or not header_key(name):
                raise ValueError("Пустой или неверный псевдоним столбца.")
            key = header_key(name)
            if key in aliases:
                raise ValueError(f"Неоднозначный псевдоним столбца: {name}")
            aliases[key] = field
    return ColumnRegistry(aliases)
````

## src/kuchenland_importer/infrastructure/capture_loader.py

````python
"""Verify a capture manifest before processing its source attachments."""

import json
from datetime import datetime
from hashlib import file_digest
from pathlib import Path

from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment


def load_attachments(manifest: Path) -> tuple[SavedAttachment, ...]:
    root = manifest.resolve().parent
    data = json.loads(manifest.read_text(encoding="utf-8"))
    if data.get("schema_version") != 1 or data.get("stage") != "mail_capture":
        raise ValueError("Неподдерживаемый манифест получения писем.")
    mails = [
        MailMetadata(
            m["entry_id"],
            m["store_id"],
            m["subject"],
            datetime.fromisoformat(m["received_at"]),
            m["sender_name"],
            m["sender_address"],
        )
        for m in data["mails"]
    ]
    result = []
    for item in data["attachments"]:
        path = (root / item["path"]).resolve()
        if not path.is_relative_to(root):
            raise ValueError("Путь вложения выходит за пределы каталога запуска.")
        with path.open("rb") as stream:
            digest = file_digest(stream, "sha256").hexdigest()
        if digest != item["sha256"]:
            raise ValueError("Контрольная сумма вложения не совпадает с манифестом.")
        index = item["mail_index"]
        if type(index) is not int or index < 0 or index >= len(mails):
            raise ValueError("Неверный индекс письма в манифесте.")
        result.append(
            SavedAttachment(
                mails[index], item["attachment_index"], item["original_name"], path, digest
            )
        )
    return tuple(result)
````

## src/kuchenland_importer/infrastructure/preview_report.py

````python
"""Serialize typed values explicitly, preserving Excel error types."""

import json
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from kuchenland_importer.application.normalize_workbook import NormalizeWorkbook
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.capture_loader import load_attachments
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader


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


def create_preview(manifest: Path, columns: Path, output: Path) -> PreviewSummary:
    try:
        attachments = load_attachments(manifest)
        service = NormalizeWorkbook(load_columns(columns))
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ExcelInputError(f"Не удалось прочитать входные данные: {error}") from error
    result: dict[str, object] = {"schema_version": 1, "stage": "excel_preview"}
    files = []
    product_count = error_count = 0
    for attachment in attachments:
        try:
            preview = service.execute(attachment, WorkbookReader().read(attachment.path))
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
            issues = [
                {
                    "code": i.code,
                    "message": i.message,
                    "severity": i.severity.value,
                    "source": i.source,
                }
                for i in preview.issues
            ]
            errors = sum(i.severity.value == "error" for i in preview.issues)
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
            error_count += errors
        except Exception as error:
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
    partial = output.with_suffix(".json.part")
    try:
        with partial.open("x", encoding="utf-8") as stream:
            json.dump(
                result, stream, default=encode_value, ensure_ascii=False, indent=2, allow_nan=False
            )
        partial.rename(output)
    except (OSError, ValueError, TypeError) as error:
        raise ExcelInputError(f"Не удалось сохранить отчёт: {error}") from error
    return PreviewSummary(output, product_count, error_count)
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
    parser = argparse.ArgumentParser(description="Kuchenland: диагностика и получение писем")
    parser.add_argument(
        "command",
        nargs="?",
        default="diagnose",
        choices=("diagnose", "capture-outlook", "preview-excel"),
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
        if args.command == "preview-excel":
            from uuid import uuid4

            from kuchenland_importer.infrastructure.preview_report import create_preview

            if args.manifest is None:
                parser.error("preview-excel требует --manifest")
            preview = create_preview(
                args.manifest,
                args.config.resolve().parent / "columns.toml",
                app.paths.reports / f"excel-preview-{uuid4().hex}.json",
            )
            print(f"Прочитано товаров: {preview.products}; ошибок: {preview.errors}")
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

## config/columns.toml

````toml
[fields]
article = ["Артикул", "Article", "SKU", "Item No", "Item Number"]
supplier = ["Поставщик", "Supplier", "Производитель", "Manufacturer"]
season = ["Сезон", "Season"]
name = ["Наименование", "Название", "Product Name", "Description of goods"]
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

## tools/extract_msg_samples.py

````python
"""QA utility for the supplied Unicode MSG samples, not a general MSG importer."""

import argparse
import hashlib
import json
import struct
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pythoncom
import pywintypes


def read_stream(storage, name: str) -> bytes:
    stream = storage.OpenStream(name, None, 0x10, 0)
    chunks = []
    while chunk := stream.Read(1024 * 1024):
        chunks.append(chunk)
    return b"".join(chunks)


def text(storage, tag: str) -> str:
    try:
        return read_stream(storage, f"__substg1.0_{tag}001F").decode("utf-16-le").rstrip("\0")
    except pywintypes.com_error:
        return ""


def delivery_time(storage) -> datetime:
    properties = read_stream(storage, "__properties_version1.0")
    for offset in range(32, len(properties) - 15, 16):
        if struct.unpack_from("<I", properties, offset)[0] == 0x0E060040:
            ticks = struct.unpack_from("<Q", properties, offset + 8)[0]
            return datetime(1601, 1, 1, tzinfo=UTC) + timedelta(microseconds=ticks // 10)
    raise ValueError("MSG sample has no delivery time; sent date cannot replace it")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    data = {
        "schema_version": 1,
        "stage": "mail_capture",
        "run_id": "msg-qa-samples",
        "selected_count": 0,
        "mails": [],
        "attachments": [],
        "issues": [],
    }
    for index, path in enumerate(sorted(args.source.glob("*.msg")), 1):
        original = path.read_bytes()
        digest = hashlib.sha256(original).hexdigest()
        storage = pythoncom.StgOpenStorage(str(path.resolve()), None, 0x20, None, 0)
        data["mails"].append(
            {
                "selection_index": index,
                "entry_id": "msg:" + digest,
                "store_id": "offline-msg-samples",
                "subject": text(storage, "0037"),
                "received_at": delivery_time(storage).isoformat(),
                "sender_name": text(storage, "0C1A"),
                "sender_address": text(storage, "5D01") or text(storage, "0C1F"),
            }
        )
        folder = output / f"mail-{index:04d}"
        folder.mkdir()
        for name, kind, *_ in storage.EnumElements():
            if kind != 1 or not name.startswith("__attach_version1.0_"):
                continue
            attachment = storage.OpenStorage(name, None, 0x10, None, 0)
            filename = text(attachment, "3707") or text(attachment, "3704")
            suffix = Path(filename).suffix.lower()
            if suffix not in {".xlsx", ".xlsm", ".xls", ".xlsb"}:
                continue
            content = read_stream(attachment, "__substg1.0_37010102")
            target = folder / f"attachment-{len(data['attachments']) + 1:04d}{suffix}"
            target.write_bytes(content)
            data["attachments"].append(
                {
                    "mail_index": index - 1,
                    "attachment_index": int(name.rsplit("#", 1)[1], 16) + 1,
                    "original_name": filename,
                    "path": target.relative_to(output).as_posix(),
                    "sha256": hashlib.sha256(content).hexdigest(),
                }
            )
        storage = None
        assert hashlib.sha256(path.read_bytes()).hexdigest() == digest
    data["selected_count"] = len(data["mails"])
    (output / "manifest.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"Extracted {len(data['attachments'])} Excel attachments from {len(data['mails'])} MSG")


if __name__ == "__main__":
    main()
````

## tests/unit/test_excel_normalization.py

````python
import json
import zipfile
from datetime import UTC, date, datetime
from decimal import Decimal
from hashlib import sha256
from pathlib import Path

import pytest
from openpyxl import Workbook

from kuchenland_importer.application.normalize_workbook import NormalizeWorkbook, article_value
from kuchenland_importer.application.value_normalization import normalize_value
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.workbook import SourceCell
from kuchenland_importer.infrastructure.capture_loader import load_attachments
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader
from kuchenland_importer.infrastructure.preview_report import encode_value

CONFIG = Path(__file__).resolve().parents[2] / "config" / "columns.toml"


def attached(path: Path) -> SavedAttachment:
    mail = MailMetadata("entry", "store", "Subject", datetime.now(UTC), "", "")
    return SavedAttachment(mail, 1, path.name, path, sha256(path.read_bytes()).hexdigest())


def preview(path: Path):
    return NormalizeWorkbook(load_columns(CONFIG)).execute(
        attached(path), WorkbookReader().read(path)
    )


def test_reordered_synonyms_unknown_fields_and_photo_join(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    workbook = Workbook()
    main = workbook.active
    main.append(["Season", "Supplier", "Unit Price", "SKU", "Custom field"])
    main.append(["Весна 2027", "Vendor", "1 234,50", "001-A", "Keep me"])
    main.append(["Лето 2027", "Vendor", 0, "B", "Other"])
    photo = workbook.create_sheet("Images")
    photo.append(["Артикул", "Фото", "Описание"])
    photo.append(["B", None, "Second"])
    photo.append(["001-A", None, "First"])
    workbook.save(path)
    before = path.read_bytes()
    result = preview(path)
    assert not result.issues
    assert len(result.products) == 2
    first = result.products[0]
    assert first.article == "001-A"
    assert first.values["purchase_price"] == Decimal("1234.50")
    assert first.values["source:custom field"] == "Keep me"
    assert first.values["photo:description"] == "First"
    assert result.products[1].values["purchase_price"] == Decimal("0")
    assert path.read_bytes() == before


def test_multilevel_header_with_merged_groups(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append(["Артикул", "Поставщик", "Сезон", "FOB", None, "Наименование"])
    sheet.append([None, None, None, "ПРЕДЫДУЩИЙ", "НОВЫЙ", None])
    for col in ("A", "B", "C", "F"):
        sheet.merge_cells(f"{col}1:{col}2")
    sheet.merge_cells("D1:E1")
    sheet.append(["A", "Vendor", "Зима", 5, 6, "Name"])
    book.save(path)
    result = preview(path)
    assert len(result.products) == 1
    assert result.products[0].values["purchase_price"] == Decimal("6")
    assert result.products[0].source_row == 3


def test_excel_error_remains_typed_and_formula_without_cache_fails(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append(["Артикул", "Поставщик", "Сезон", "Цена"])
    sheet.append(["A", "Vendor", "Весна", "#DIV/0!"])
    sheet.append(["B", "Vendor", "Весна", "=1+1"])
    book.save(path)
    result = preview(path)
    assert len(result.products) == 1
    assert result.products[0].values["purchase_price"] == ExcelErrorValue("#DIV/0!")
    assert {issue.code for issue in result.issues} == {"EXCEL_ERROR_VALUE", "ROW_INVALID"}
    assert encode_value(ExcelErrorValue("#DIV/0!")) == {"type": "excel_error", "code": "#DIV/0!"}


def test_cached_empty_string_is_not_missing_formula_cache(tmp_path: Path) -> None:
    path = tmp_path / "empty-result.xlsx"
    book = Workbook()
    book.active.append(["Артикул", "Поставщик", "Сезон", "Optional"])
    book.active.append(["A", "Vendor", "Весна", '=IF(1=1,"","")'])
    book.save(path)
    # Produce the OOXML cached empty-string representation saved by Excel.
    with zipfile.ZipFile(path) as archive:
        contents = {name: archive.read(name) for name in archive.namelist()}
    contents["xl/worksheets/sheet1.xml"] = contents["xl/worksheets/sheet1.xml"].replace(
        b'<c r="D2">', b'<c r="D2" t="str">'
    )
    with zipfile.ZipFile(path, "w") as archive:
        for name, value in contents.items():
            archive.writestr(name, value)
    result = preview(path)
    assert len(result.products) == 1
    assert not result.issues
    assert result.products[0].values["source:optional"] == ""


@pytest.mark.parametrize("text", ["SKU", "Article", "Артикул", "Item No."])
def test_article_aliases(text: str) -> None:
    assert load_columns(CONFIG).resolve(text) == "article"


def test_leading_zeros_and_no_fractional_article() -> None:
    assert article_value(SourceCell(12, "00000")) == "00012"
    assert article_value(SourceCell("00012")) == "00012"
    with pytest.raises(ValueError):
        article_value(SourceCell(12.5))


def test_ambiguous_sku_columns_are_rejected(tmp_path: Path) -> None:
    path = tmp_path / "bad.xlsx"
    book = Workbook()
    book.active.append(["Артикул", "SKU", "Поставщик", "Сезон"])
    book.active.append(["A", "B", "Vendor", "Весна"])
    book.save(path)
    result = preview(path)
    assert not result.products
    assert any(issue.code == "SHEET_SCHEMA" for issue in result.issues)


def test_no_supplier_guessing_from_sender(tmp_path: Path) -> None:
    path = tmp_path / "no-vendor.xlsx"
    book = Workbook()
    book.active.append(["Артикул", "Поставщик", "Сезон"])
    book.active.append(["A", None, "Весна"])
    book.save(path)
    assert not preview(path).products


def test_manifest_path_cannot_escape_capture_directory(tmp_path: Path) -> None:
    path = tmp_path / "manifest.json"
    path.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "stage": "mail_capture",
                "mails": [],
                "attachments": [{"path": "../outside.xlsx"}],
            }
        ),
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="выходит"):
        load_attachments(path)


def test_date_normalization_without_inventing_year() -> None:
    assert normalize_value("sales_start", "03.01.2027") == date(2027, 1, 3)
    assert normalize_value("sales_start", "декабрь") == "декабрь"
````

## tests/unit/test_legacy_reader.py

````python
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from kuchenland_importer.infrastructure.excel import legacy_reader


def test_owned_excel_is_closed_after_open_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    app = SimpleNamespace(
        Workbooks=SimpleNamespace(Open=Mock(side_effect=RuntimeError("failed"))), Quit=Mock()
    )
    initialize, uninitialize = Mock(), Mock()
    monkeypatch.setattr(legacy_reader.pythoncom, "CoInitializeEx", initialize)
    monkeypatch.setattr(legacy_reader.pythoncom, "CoUninitialize", uninitialize)
    monkeypatch.setattr(legacy_reader.win32com.client, "DispatchEx", Mock(return_value=app))
    with pytest.raises(RuntimeError, match="failed"):
        legacy_reader.read_legacy(tmp_path / "source.xls")
    assert app.AutomationSecurity == 3
    assert app.EnableEvents is False
    assert app.Workbooks.Open.call_args.kwargs["UpdateLinks"] == 0
    assert app.Workbooks.Open.call_args.kwargs["ReadOnly"] is True
    app.Quit.assert_called_once()
    uninitialize.assert_called_once()


def test_legacy_book_never_saved(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    book = SimpleNamespace(Sheets=[], Worksheets=[], Close=Mock())
    app = SimpleNamespace(Workbooks=SimpleNamespace(Open=Mock(return_value=book)), Quit=Mock())
    monkeypatch.setattr(legacy_reader.pythoncom, "CoInitializeEx", Mock())
    monkeypatch.setattr(legacy_reader.pythoncom, "CoUninitialize", Mock())
    monkeypatch.setattr(legacy_reader.win32com.client, "DispatchEx", Mock(return_value=app))
    assert legacy_reader.read_legacy(tmp_path / "source.xlsb") == ()
    book.Close.assert_called_once_with(SaveChanges=False)
    app.Quit.assert_called_once()
````

## tests/unit/test_models.py

````python
from datetime import UTC, datetime
from decimal import Decimal
from pathlib import Path
from typing import cast

import pytest

from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.product import CellValue, ProductRecord
from kuchenland_importer.domain.report import ImportIssue, ImportReport, Severity


def attachment(tmp_path: Path) -> SavedAttachment:
    mail = MailMetadata("entry", "store", "Тема", datetime(2026, 10, 9, tzinfo=UTC), "Закупщик", "")
    return SavedAttachment(mail, 1, "товары.xlsx", tmp_path / "товары.xlsx", "a" * 64)


def test_article_zeros_and_defensive_copy(tmp_path: Path) -> None:
    values: dict[str, CellValue] = {"price": Decimal("15.20")}
    product = ProductRecord(
        " 001-Ab ", "Поставщик", "Весна 2027", attachment(tmp_path), "Товары", 2, values
    )
    values["price"] = Decimal("999")
    assert product.article == "001-Ab"
    assert product.values["price"] == Decimal("15.20")
    with pytest.raises(TypeError):
        cast(dict[str, CellValue], product.values)["price"] = 0


@pytest.mark.parametrize("value", [float("nan"), float("inf"), Decimal("NaN")])
def test_non_finite_values_rejected(tmp_path: Path, value: CellValue) -> None:
    with pytest.raises(ValueError):
        ProductRecord(
            "001", "Поставщик", "Весна", attachment(tmp_path), "Лист", 2, {"price": value}
        )


def test_received_date_requires_timezone() -> None:
    with pytest.raises(ValueError, match="часовой пояс"):
        MailMetadata("entry", "store", "", datetime(2026, 10, 9), "", "")


def test_report_separates_errors_and_warnings() -> None:
    report = ImportReport(
        "run",
        issues=(
            ImportIssue("NO_PHOTO", "Фото отсутствует", Severity.WARNING, "book.xlsx:2"),
            ImportIssue("NO_SKU", "Нет артикула", Severity.ERROR, "book.xlsx:3"),
        ),
    )
    assert report.error_count == 1
    assert report.warning_count == 1


def test_report_rejects_negative_count() -> None:
    with pytest.raises(ValueError):
        ImportReport("run", products_added=-1)
````

## tests/integration/test_excel_preview.py

````python
import json
from datetime import UTC, datetime
from hashlib import sha256
from pathlib import Path

import pytest
from openpyxl import Workbook

from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.preview_report import create_preview

CONFIG = Path(__file__).resolve().parents[2] / "config" / "columns.toml"


def manifest(tmp_path: Path) -> Path:
    good = tmp_path / "good.xlsx"
    book = Workbook()
    book.active.append(["Артикул", "Поставщик", "Сезон"])
    book.active.append(["A", "Vendor", "Весна"])
    book.save(good)
    bad = tmp_path / "broken.xlsx"
    bad.write_bytes(b"not Excel")
    path = tmp_path / "manifest.json"
    path.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "stage": "mail_capture",
                "mails": [
                    {
                        "entry_id": "entry",
                        "store_id": "store",
                        "subject": "",
                        "sender_name": "",
                        "sender_address": "",
                        "received_at": datetime.now(UTC).isoformat(),
                    }
                ],
                "attachments": [
                    {
                        "mail_index": 0,
                        "attachment_index": index,
                        "original_name": p.name,
                        "path": p.name,
                        "sha256": sha256(p.read_bytes()).hexdigest(),
                    }
                    for index, p in enumerate((good, bad), 1)
                ],
            }
        ),
        encoding="utf-8",
    )
    return path


def test_bad_workbook_does_not_block_good_attachment(tmp_path: Path) -> None:
    source = manifest(tmp_path)
    result = create_preview(source, CONFIG, tmp_path / "preview.json")
    assert result.products == 1
    assert result.errors == 1
    data = json.loads(result.report.read_text(encoding="utf-8"))
    assert data["files"][0]["eligible_for_import"] is True
    assert data["files"][1]["eligible_for_import"] is False
    assert not list(tmp_path.glob("*.part"))


def test_modified_source_is_detected_before_processing(tmp_path: Path) -> None:
    source = manifest(tmp_path)
    (tmp_path / "good.xlsx").write_bytes(b"changed")
    with pytest.raises(ExcelInputError, match="сумма"):
        create_preview(source, CONFIG, tmp_path / "preview.json")
    assert not (tmp_path / "preview.json").exists()
````

## pyproject.toml

````toml
[build-system]
requires = ["setuptools>=75,<83"]
build-backend = "setuptools.build_meta"

[project]
name = "kuchenland-importer"
version = "0.3.0"
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

## README.md

````markdown
# Kuchenland Importer

Windows-приложение для импорта товаров из выделенных писем классического Outlook
в локальную общую книгу Excel. Python 3.12, Office Professional Plus 2024 / Microsoft 365.

## Текущий статус

Реализованы этапы 1–3: основа, получение писем и предварительный разбор Excel.
На реальных XLSX-вложениях из 13 MSG прочитано 90 товаров. Автоматические тесты проверены;
приёмка на настоящем Outlook отложена до устройства с настроенной учётной записью.
Импорт товаров в Excel, работа с фото, GUI и EXE ещё не реализованы.
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
по артикулу и сохраняются с префиксом photo:, изображения пока не извлекаются.
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
Реальные вложения проверены для XLSX; испытания XLS/XLSB через Office ещё впереди.
Файлы с паролем не обещаются к поддержке без отдельного решения.

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

Рабочие письма и книги, логи, резервные копии, `.venv` и личные настройки исключены
из Git. После каждого завершённого этапа и успешных проверок код публикуется
в [MailExcel2Git](https://github.com/MaxGarAI/MailExcel2Git) в ветку
`codex/development`. Порядок публикации закреплён в `AGENTS.md`.
````

## docs/specification.md

````markdown
# Спецификация импорта

Версия 0.1. Уточнения пользователя от 09.10.2026 (Europe/Moscow).

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
- Год сезона при декабрьском переходе; строка «Зима» без начала продаж;
  составные сезоны вне утверждённых псевдонимов.
- Название поля начала продаж и поставщика в реальных вложениях.
- Что делать с ручными значениями в столбцах, отсутствующих во вложении:
  пользователь явно согласовал сохранение формул, но не всех ручных значений.
- Шаблон новой расчётной/фото-пары; какие формулы и оформление переносить.
- Число пользователей и работа с открытой/несохранённой книгой. До согласования
  запись в открытую пользователем книгу не включается.
- Поведение для нескольких фото на один артикул, группированных изображений,
  повреждённых/защищённых файлов и формул без сохранённых результатов.
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
Загрузка конфигурации ещё не означает реализацию выбора сезона (этап 4).

## Модели и данные

Артикул — строка, ведущие нули и регистр сохраняются. Текстовые поля не преобразуются
в формулы Excel. ProductRecord содержит копию словаря значений с запретом изменений,
ссылку на письмо/вложение и координаты исходной строки.
Фото имеет хеш нормализованных пикселей; фактическое извлечение относится к этапу 5.
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
````

## docs/stage-3-validation.md

````markdown
# Этап 3: чтение и нормализация Excel

09.10.2026, Python 3.12.8, Windows, версия приложения 0.3.0.

- Полный набор pytest: 61 passed.
- Ruff: All checks passed.
- Mypy strict: Success, 36 source files.
- Вложения извлечены из всех 13 предоставленных Unicode MSG через Windows IStorage
  без Outlook. Исходные MSG проверены по SHA-256; исходные письма не изменялись.
- На 13 реальных XLSX-вложениях получены 90 товаров, ошибок нормализации 0.
  Сохранены 32 значения ошибок ячеек Excel; они отражены предупреждениями.
- Проверены перестановка колонок, синонимы, неизвестные поля, фото-описания по ключу,
  объединённые заголовки, нули артикулов, деньги, даты, пустой кеш строковой формулы,
  отсутствующий кеш, дубли обязательных столбцов и отсутствие поставщика.
- Проверены повреждённая книга, изменённая контрольная сумма, путь вне каталога,
  продолжение обработки других вложений, VBA-параметры и закрытие owned Excel на тестовых
  объектах. Реальные XLS/XLSB и COM-чтение на Office 2024/365 НЕ проверены.

OOXML читается через openpyxl, не сохраняется этим адаптером. Предупреждение о
неподдерживаемом расширении условного форматирования относится к модели чтения,
а не к изменению исходника. SHA-256 до/после совпал для всех 27 файлов:
13 извлечённых вложений, 13 исходных MSG и общей книги. Финальный локальный отчёт:
`.runtime/reports/excel-preview-8ea433a9ab934d0d8918c9a7fb75386c.json`.

Предпросмотр не выбирает победителя между письмами, не определяет сезонную вкладку,
не переносит фотографии и не записывает общую книгу. Это следующие этапы.
Без сохранённого значения формулы автоматического расчёта в OOXML-адаптере нет;
нужен расчёт в Excel и повторное получение вложения. Зашифрованные источники,
Excel 4.0 macros и произвольные нестандартные шаблоны не заявлены как проверенные.
Пределы и порядок запуска описаны в README.

[Контракт data_only и сохранённых значений формул](https://openpyxl.readthedocs.io/en/3.1/tutorial.html).
[Контракт ReadOnly, UpdateLinks и AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open).
````
