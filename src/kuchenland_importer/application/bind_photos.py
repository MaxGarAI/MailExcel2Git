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
