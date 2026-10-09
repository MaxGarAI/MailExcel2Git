"""Extensible seasonal destinations and traceable routing decisions."""

import re
import unicodedata
from dataclasses import dataclass

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.report import ImportIssue


def normalize_alias(value: str) -> str:
    text = unicodedata.normalize("NFKC", value).casefold().replace("ё", "е")
    text = re.sub(r"[‐‑‒–—−]", "-", text)
    return re.sub(r"\s*-\s*", "-", " ".join(text.split()))


def validate_sheet_name(name: str) -> None:
    if (
        not name.strip()
        or name != name.strip()
        or len(name) > 31
        or any(character in name for character in "[]:*?/\\")
        or any(ord(character) < 32 for character in name)
        or name.startswith("'")
        or name.endswith("'")
    ):
        raise ConfigurationError(f"Недопустимое имя вкладки Excel: {name!r}.")


@dataclass(frozen=True, slots=True)
class SeasonRoute:
    id: str
    aliases: tuple[str, ...]
    calculation_sheet: str
    photo_sheet: str

    def __post_init__(self) -> None:
        if not re.fullmatch(r"[a-z][a-z0-9_]*", self.id):
            raise ConfigurationError(f"Недопустимый идентификатор сезона: {self.id!r}.")
        if not self.aliases or any(not alias.strip() for alias in self.aliases):
            raise ConfigurationError(f"Сезону {self.id} нужны непустые названия.")
        validate_sheet_name(self.calculation_sheet)
        validate_sheet_name(self.photo_sheet)


@dataclass(frozen=True, slots=True)
class SeasonAssignment:
    source_season: str
    source_year: int | None
    effective_year: int | None
    effective_season: str
    route: SeasonRoute
    sales_start_month: int | None
    rule: str
    issues: tuple[ImportIssue, ...] = ()


@dataclass(frozen=True, slots=True)
class RoutedProduct:
    product: ProductRecord
    assignment: SeasonAssignment | None


@dataclass(frozen=True, slots=True)
class RoutedWorkbook:
    products: tuple[RoutedProduct, ...]
    issues: tuple[ImportIssue, ...]
