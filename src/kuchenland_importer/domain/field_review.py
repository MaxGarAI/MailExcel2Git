"""Explicit, supplier-scoped decisions for fields missing in a destination layout."""

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class FieldChoice:
    supplier: str
    photo: bool
    source: str
    samples: tuple[str, ...]
    targets: tuple[str, ...]
    sheets: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class FieldDecision:
    action: str  # skip, map, create
    target: str = ""


@dataclass(frozen=True, slots=True)
class AddedColumn:
    sheet: str
    key: str
    label: str
    column: int


def rule_key(workbook: Path, supplier: str, photo: bool, source: str) -> str:
    return json.dumps([str(workbook.resolve()).casefold(), supplier.strip().casefold(),
                       photo, source], ensure_ascii=False)


