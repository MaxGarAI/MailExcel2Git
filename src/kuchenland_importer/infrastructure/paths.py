"""Create and probe writable application directories without touching workbooks."""

from dataclasses import dataclass
from pathlib import Path
from typing import Self
from uuid import uuid4

from kuchenland_importer.domain.errors import StartupError


@dataclass(frozen=True, slots=True)
class RuntimePaths:
    logs: Path
    backup: Path
    work: Path
    reports: Path

    @classmethod
    def prepare(cls, root: Path) -> Self:
        paths = cls(*(root / name for name in ("logs", "backup", "work", "reports")))
        for directory in (paths.logs, paths.backup, paths.work, paths.reports):
            probe = directory / f".write-check-{uuid4().hex}"
            try:
                directory.mkdir(parents=True, exist_ok=True)
                with probe.open("xb") as stream:
                    stream.write(b"write-check")
                probe.unlink()
            except OSError as error:
                raise StartupError(f"Рабочий каталог недоступен: {directory}: {error}") from error
        return paths
