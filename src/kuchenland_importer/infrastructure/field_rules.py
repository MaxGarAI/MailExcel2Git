"""Atomic local rule storage; no workbook data or sample values are persisted."""

import json
import os
from pathlib import Path
from uuid import uuid4

from kuchenland_importer.domain.field_review import FieldDecision, rule_key  # noqa: F401


def load_rules(path: Path) -> dict[str, FieldDecision]:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schema_version") != 1 or not isinstance(data.get("rules"), dict):
        raise ValueError("Неизвестный формат правил дополнительных полей.")
    result = {}
    for key, value in data["rules"].items():
        decision = FieldDecision(**value)
        if decision.action not in {"skip", "map", "create"}:
            raise ValueError("Недопустимое действие правила поля.")
        result[key] = decision
    return result


def save_rules(path: Path, rules: dict[str, FieldDecision]) -> None:
    from dataclasses import asdict

    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}-{uuid4().hex}.tmp")
    try:
        with temporary.open("x", encoding="utf-8") as stream:
            json.dump({"schema_version": 1, "rules": {k: asdict(v) for k, v in rules.items()}},
                      stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
