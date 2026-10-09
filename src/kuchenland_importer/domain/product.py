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
