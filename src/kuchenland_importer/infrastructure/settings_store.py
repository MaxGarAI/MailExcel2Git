"""Validated atomic TOML persistence for personal settings."""

import json
import os
from pathlib import Path
from uuid import uuid4

from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.settings import Settings


def save_settings(path: Path, settings: Settings) -> None:
    def quote(value: str) -> str:
        return json.dumps(value, ensure_ascii=False)

    def array(values: tuple[str, ...]) -> str:
        return "[" + ", ".join(quote(value) for value in values) + "]"

    text = (
        "schema_version = 1\n\n[application]\n"
        f"workbook = {quote(str(settings.workbook) if settings.workbook else '')}\n"
        f"data_dir = {quote(str(settings.data_dir))}\n\n[logging]\n"
        f"level = {quote(settings.log_level)}\nrotation_mb = {settings.log_rotation_mb}\n"
        f"retention_days = {settings.log_retention_days}\n\n[excel]\n"
        f"input_extensions = {array(settings.input_extensions)}\n"
        f"excluded_sheets = {array(settings.excluded_sheets)}\n"
    )
    for route in settings.routes:
        text += (
            f"\n[[routes]]\nid = {quote(route.id)}\naliases = {array(route.aliases)}\n"
            f"calculation_sheet = {quote(route.calculation_sheet)}\n"
            f"photo_sheet = {quote(route.photo_sheet)}\n"
        )
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}-{uuid4().hex}.tmp")
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as stream:
            stream.write(text)
            stream.flush()
            os.fsync(stream.fileno())
        if load_settings(temporary) != settings:
            raise ValueError("Настройки не прошли проверку после сохранения.")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
