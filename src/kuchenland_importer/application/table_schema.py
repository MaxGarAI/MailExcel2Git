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
