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
