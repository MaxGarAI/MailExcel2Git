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
