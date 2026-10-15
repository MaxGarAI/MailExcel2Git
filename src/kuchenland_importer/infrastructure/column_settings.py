"""Persist personal column aliases after validating all ambiguity rules."""

import json
import os
import tomllib
from pathlib import Path
from uuid import uuid4

from kuchenland_importer.infrastructure.excel.columns import load_columns


def column_path(config: Path) -> Path:
    personal = config.parent / "columns.local.toml"
    return personal if personal.exists() else config.parent / "columns.toml"


def read_fields(path: Path) -> dict[str, list[str]]:
    load_columns(path)
    with path.open("rb") as stream:
        fields: dict[str, list[str]] = tomllib.load(stream)["fields"]
    return fields


def save_fields(path: Path, fields: dict[str, list[str]]) -> None:
    if not {"article", "supplier", "season", "photo"} <= fields.keys():
        raise ValueError("Обязательные поля: article, supplier, season, photo.")
    quote = json.dumps
    content = "[fields]\n" + "".join(
        f"{quote(key)} = {json.dumps(values, ensure_ascii=False)}\n"
        for key, values in fields.items()
    )
    temporary = path.with_name(f".{path.name}-{uuid4().hex}.tmp")
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        load_columns(temporary)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
