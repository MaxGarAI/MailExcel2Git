"""Explicit write policy, separated from mail/source-reader configuration."""

import tomllib
from dataclasses import dataclass
from pathlib import Path

from kuchenland_importer.domain.errors import ConfigurationError


@dataclass(frozen=True, slots=True)
class WritingSettings:
    preserve_manual_values: bool
    unknown_fields: str
    missing_sheets: str


def load_writing_settings(path: Path) -> WritingSettings:
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
        expected = {
            "schema_version",
            "preserve_manual_values",
            "unknown_fields",
            "missing_sheets",
        }
        if set(data) != expected or type(data["schema_version"]) is not int:
            raise ValueError("Неверные поля настроек записи.")
        if data["schema_version"] != 1:
            raise ValueError("Поддерживается schema_version=1.")
        if type(data["preserve_manual_values"]) is not bool:
            raise ValueError("preserve_manual_values должно быть логическим.")
        if data["unknown_fields"] != "skip" or data["missing_sheets"] != "skip_attachment":
            raise ValueError("Поддерживаются unknown_fields=skip и missing_sheets=skip_attachment.")
        return WritingSettings(
            data["preserve_manual_values"],
            data["unknown_fields"],
            data["missing_sheets"],
        )
    except (OSError, ValueError) as error:
        raise ConfigurationError(f"Не удалось прочитать настройки записи: {error}") from error
