"""Persist a result report without turning an already committed workbook into a false failure."""

import json
from pathlib import Path

from loguru import logger


def save_report(path: Path, data: dict[str, object], *, committed: bool) -> bool:
    try:
        partial = path.with_suffix(".part")
        with partial.open("x", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
        partial.replace(path)
        return True
    except (OSError, TypeError, ValueError):
        logger.exception("Книга сохранена={}, отчёт недоступен: {}", committed, path)
        return False
