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
