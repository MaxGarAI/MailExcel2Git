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
