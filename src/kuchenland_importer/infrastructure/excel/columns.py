"""Load the configurable field dictionary without embedding IO in business rules."""

import tomllib
from pathlib import Path

from kuchenland_importer.application.table_schema import ColumnRegistry, header_key


def load_columns(path: Path) -> ColumnRegistry:
    with path.open("rb") as stream:
        data = tomllib.load(stream)
    if set(data) != {"fields"} or not isinstance(data["fields"], dict):
        raise ValueError("Ожидается таблица fields в настройках столбцов.")
    aliases = {}
    for field, names in data["fields"].items():
        if not isinstance(names, list) or not names:
            raise ValueError("Поле должно иметь список псевдонимов.")
        for name in names:
            if not isinstance(name, str) or not header_key(name):
                raise ValueError("Пустой или неверный псевдоним столбца.")
            key = header_key(name)
            if key in aliases:
                raise ValueError(f"Неоднозначный псевдоним столбца: {name}")
            aliases[key] = field
    return ColumnRegistry(aliases)
