"""Validated immutable configuration; seasons are extensible data, not an enum."""

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

from kuchenland_importer.domain.errors import ConfigurationError


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
class Settings:
    workbook: Path | None
    data_dir: Path
    log_level: str
    log_rotation_mb: int
    log_retention_days: int
    routes: tuple[SeasonRoute, ...]
    excluded_sheets: tuple[str, ...]
    input_extensions: tuple[str, ...]

    def __post_init__(self) -> None:
        if self.workbook is not None and not self.workbook.is_absolute():
            raise ConfigurationError("Путь общей книги должен быть абсолютным.")
        if not self.data_dir.is_absolute():
            raise ConfigurationError("Путь рабочего каталога должен быть абсолютным.")
        if self.log_level not in {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}:
            raise ConfigurationError("Неизвестный уровень логирования.")
        if self.log_rotation_mb < 1 or self.log_retention_days < 1:
            raise ConfigurationError("Ротация и срок хранения логов должны быть положительными.")
        allowed = {".xlsx", ".xlsm", ".xls", ".xlsb"}
        if not self.input_extensions or not set(self.input_extensions) <= allowed:
            raise ConfigurationError("Разрешены форматы .xlsx, .xlsm, .xls, .xlsb.")
        if len(set(self.input_extensions)) != len(self.input_extensions):
            raise ConfigurationError("Расширения вложений не должны повторяться.")
        if not self.routes:
            raise ConfigurationError("Нужно определить хотя бы один сезон.")
        ids: set[str] = set()
        aliases: set[str] = set()
        sheets: set[str] = set()
        excluded = {name.casefold() for name in self.excluded_sheets}
        for route in self.routes:
            if route.id in ids:
                raise ConfigurationError(f"Сезон {route.id} указан повторно.")
            ids.add(route.id)
            for alias in route.aliases:
                normalized = normalize_alias(alias)
                if normalized in aliases:
                    raise ConfigurationError(f"Название сезона неоднозначно: {alias!r}.")
                aliases.add(normalized)
            for sheet in (route.calculation_sheet, route.photo_sheet):
                key = sheet.casefold()
                if key in sheets or key in excluded:
                    raise ConfigurationError(f"Вкладка {sheet!r} назначена неоднозначно.")
                sheets.add(key)
