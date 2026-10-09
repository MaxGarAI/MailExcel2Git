"""Master sheet layouts and existing row identities; articles remain case-sensitive text."""

from dataclasses import dataclass
from datetime import date, datetime

from kuchenland_importer.domain.product import PhotoAsset


@dataclass(slots=True)
class SheetLayout:
    name: str
    header_row: int
    header_depth: int
    first_row: int
    columns: dict[str, int]
    last_row: int
    calculation_sheet: str | None = None


@dataclass(frozen=True, slots=True)
class ExistingRow:
    sheet: str
    row: int


@dataclass(slots=True)
class ExistingProduct:
    rows: list[ExistingRow]
    photo: PhotoAsset | None = None
    photo_error: str | None = None
    received_at: datetime | None = None
    fingerprint: str | None = None
    received_day: date | None = None


@dataclass(slots=True)
class MasterSnapshot:
    layouts: dict[str, SheetLayout]
    products: dict[str, ExistingProduct]
    sheet_names: tuple[str, ...]
    calculation_mode: str = "auto"
