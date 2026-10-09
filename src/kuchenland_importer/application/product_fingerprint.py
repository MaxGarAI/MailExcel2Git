"""Canonical typed-value fingerprints, independent of Excel and report serialization."""

import hashlib
import json
from datetime import date, datetime
from decimal import Decimal

from kuchenland_importer.domain.cell_values import ExcelErrorValue


def _encode(value: object) -> object:
    if isinstance(value, ExcelErrorValue):
        return {"type": "excel_error", "code": value.code}
    if isinstance(value, datetime | date):
        return {"type": "date", "value": value.isoformat()}
    if isinstance(value, Decimal):
        return {"type": "decimal", "value": str(value)}
    raise TypeError(f"Нельзя вычислить хеш значения {type(value).__name__}.")


def fingerprint_json(data: object) -> str:
    serialized = json.dumps(
        data,
        default=_encode,
        ensure_ascii=False,
        sort_keys=True,
        allow_nan=False,
        separators=(",", ":"),
    )
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()
