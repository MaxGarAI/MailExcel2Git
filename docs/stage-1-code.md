# Полные файлы этапа 1

Этот документ содержит полные тексты файлов для ревью. Рабочие исходники находятся в src, tests и config. Документ не импортируется приложением.

## .gitignore

```text
.venv/
.tmp/
.runtime/
__pycache__/
*.py[cod]
*.egg-info/
.pytest_cache/
.mypy_cache/
.ruff_cache/
.coverage
htmlcov/
build/
dist/
logs/*
!logs/.gitkeep
backup/*
!backup/.gitkeep
config/local.toml
*.msg
*.eml
*.xls
*.xlsx
*.xlsm
*.xlsb
~$*
```

## pyproject.toml

```toml
[build-system]
requires = ["setuptools>=75,<83"]
build-backend = "setuptools.build_meta"

[project]
name = "kuchenland-importer"
version = "0.1.0"
description = "Windows procurement mail and Excel import tool"
requires-python = ">=3.12,<3.13"
dynamic = ["dependencies"]

[project.scripts]
kuchenland-importer = "kuchenland_importer.presentation.cli:main"

[tool.setuptools.packages.find]
where = ["src"]

[tool.setuptools.dynamic]
dependencies = {file = ["requirements.txt"]}

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
addopts = "-ra --strict-markers"

[tool.ruff]
target-version = "py312"
line-length = 100

[tool.ruff.lint]
select = ["E", "F", "I", "UP", "B"]

[tool.mypy]
python_version = "3.12"
mypy_path = "src"
files = ["src"]
strict = true
```

## requirements.txt

```text
# Runtime dependencies; tested exact versions are recorded in requirements-lock.txt.
pywin32>=306; sys_platform == "win32"
xlwings>=0.33,<1
openpyxl>=3.1.5,<4
pandas>=2.2,<4
Pillow>=10.4,<13
loguru>=0.7.2,<1
# tkinter ships with the standard Windows Python distribution.
```

## requirements-dev.txt

```text
-r requirements.txt
pytest>=8.3,<10
ruff>=0.6,<1
mypy>=1.11,<2
```

## main.py

```python
"""Launch the application from a source checkout without installation."""

import sys
from pathlib import Path

if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
    from kuchenland_importer.presentation.cli import main

    raise SystemExit(main())
```

## config/default.toml

```toml
schema_version = 1

[application]
# Relative paths are resolved from this configuration file, not the current directory.
# Set the selected local workbook in config/local.toml or with --workbook.
workbook = ""
data_dir = "../.runtime"

[logging]
level = "INFO"
rotation_mb = 10
retention_days = 30

[excel]
input_extensions = [".xlsx", ".xlsm", ".xls", ".xlsb"]
excluded_sheets = ["удалено"]

[[routes]]
id = "spring"
aliases = ["Весна", "Весна-лето"]
calculation_sheet = "1. Весна 2.Весна-лето"
photo_sheet = "ФОТО весна"

[[routes]]
id = "easter"
aliases = ["Пасха"]
calculation_sheet = "8. Пасха"
photo_sheet = "ФОТО пасха"

[[routes]]
id = "summer"
aliases = ["Лето", "Лето-осень"]
calculation_sheet = "3. ЛЕТО 10.Лето-осень"
photo_sheet = "ФОТО лето"

[[routes]]
id = "autumn"
aliases = ["Осень"]
calculation_sheet = "4. ОСЕНЬ"
photo_sheet = "ФОТО осень"

[[routes]]
id = "winter"
aliases = ["Зима"]
# New sheet names; these sheets do not exist in the supplied workbook.
calculation_sheet = "Зима"
photo_sheet = "ФОТО зима"
```

## src/kuchenland_importer/__init__.py

```python
"""Kuchenland procurement importer."""

__version__ = "0.1.0"
```

## src/kuchenland_importer/app.py

```python
"""Composition root and completed stage-one diagnostics."""

from dataclasses import dataclass
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Self

from loguru import logger

from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.logging import LoggingSession, configure_logging
from kuchenland_importer.infrastructure.paths import RuntimePaths
from kuchenland_importer.infrastructure.settings import Settings


@dataclass(frozen=True, slots=True)
class DiagnosticResult:
    checks: tuple[str, ...]
    warnings: tuple[str, ...]
    errors: tuple[str, ...]


@dataclass(slots=True)
class Application:
    settings: Settings
    paths: RuntimePaths
    logging: LoggingSession

    @classmethod
    def create(
        cls, config_path: Path, *, workbook: Path | None = None, data_dir: Path | None = None
    ) -> Self:
        settings = load_settings(config_path, workbook=workbook, data_dir=data_dir)
        paths = RuntimePaths.prepare(settings.data_dir)
        session = configure_logging(paths.logs, settings)
        logger.info("Приложение запущено; настроено сезонов: {}", len(settings.routes))
        return cls(settings, paths, session)

    def diagnose(self) -> DiagnosticResult:
        checks = ["Настройки корректны.", "Рабочие каталоги доступны для записи."]
        warnings: list[str] = []
        errors: list[str] = []
        for package in ("pywin32", "xlwings", "openpyxl", "pandas", "Pillow", "loguru"):
            try:
                checks.append(f"{package}: {version(package)}")
            except PackageNotFoundError:
                errors.append(f"Не установлена зависимость {package}; установите requirements.txt.")
        workbook = self.settings.workbook
        if workbook is None:
            warnings.append("Общий файл не выбран; задайте application.workbook или --workbook.")
        else:
            try:
                if not workbook.is_file():
                    errors.append(f"Общий файл не найден: {workbook}")
                elif workbook.suffix.lower() not in self.settings.input_extensions:
                    errors.append(f"Неподдерживаемое расширение общего файла: {workbook.suffix}")
                else:
                    with workbook.open("rb") as stream:
                        stream.read(1)
                    checks.append("Общий файл доступен для чтения; запись не проверялась.")
            except OSError as error:
                errors.append(f"Общий файл недоступен: {error}")
        warnings.append("Подключение к Office и импорт будут добавлены на следующих этапах.")
        logger.info("Диагностика: ошибок {}, предупреждений {}", len(errors), len(warnings))
        return DiagnosticResult(tuple(checks), tuple(warnings), tuple(errors))

    def close(self) -> None:
        logger.info("Приложение завершено.")
        self.logging.close()
```

## src/kuchenland_importer/application/__init__.py

```python
"""Boundary for import use cases; implementation begins with the Outlook stage."""
```

## src/kuchenland_importer/domain/__init__.py

```python
"""Office-independent business models and invariants."""
```

## src/kuchenland_importer/domain/errors.py

```python
"""Errors that can be translated into actionable user messages."""


class ApplicationError(Exception):
    """Base class for expected application failures."""


class ConfigurationError(ApplicationError):
    """Configuration is missing, inconsistent, or malformed."""


class StartupError(ApplicationError):
    """Application resources cannot be initialized."""
```

## src/kuchenland_importer/domain/mail.py

```python
"""Mail provenance preserved independently of the Outlook COM object."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True, slots=True)
class MailMetadata:
    entry_id: str
    store_id: str
    subject: str
    received_at: datetime
    sender_name: str
    sender_address: str

    def __post_init__(self) -> None:
        if not self.entry_id.strip() or not self.store_id.strip():
            raise ValueError("Письмо должно иметь EntryID и StoreID.")
        if self.received_at.tzinfo is None or self.received_at.utcoffset() is None:
            raise ValueError("Дата получения письма должна содержать часовой пояс.")


@dataclass(frozen=True, slots=True)
class SavedAttachment:
    mail: MailMetadata
    attachment_index: int
    original_name: str
    path: Path
    sha256: str

    def __post_init__(self) -> None:
        if self.attachment_index < 1:
            raise ValueError("Индекс вложения Outlook начинается с 1.")
        if not self.original_name.strip() or not self.path.is_absolute():
            raise ValueError("Вложение должно иметь имя и абсолютный путь.")
        if len(self.sha256) != 64 or any(c not in "0123456789abcdef" for c in self.sha256):
            raise ValueError("Ожидается SHA-256 вложения в нижнем регистре.")
```

## src/kuchenland_importer/domain/product.py

```python
"""A product row and its provenance; source formulas are not part of this model."""

import math
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from types import MappingProxyType

from kuchenland_importer.domain.mail import SavedAttachment

type CellValue = str | int | float | bool | Decimal | date | datetime | None


@dataclass(frozen=True, slots=True)
class PhotoAsset:
    path: Path
    pixel_sha256: str
    width: int
    height: int

    def __post_init__(self) -> None:
        if not self.path.is_absolute() or self.width < 1 or self.height < 1:
            raise ValueError("Фотография должна иметь абсолютный путь и положительный размер.")
        digest = self.pixel_sha256
        if len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest):
            raise ValueError("Ожидается SHA-256 нормализованных пикселей.")


@dataclass(frozen=True, slots=True)
class ProductRecord:
    article: str
    supplier: str
    season: str
    attachment: SavedAttachment
    source_sheet: str
    source_row: int
    values: Mapping[str, CellValue]
    photo: PhotoAsset | None = None

    def __post_init__(self) -> None:
        for name in ("article", "supplier", "season", "source_sheet"):
            value = getattr(self, name)
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"Поле {name} должно быть непустой строкой.")
        if self.source_row < 1:
            raise ValueError("Номер строки Excel начинается с 1.")
        copied = dict(self.values)
        for key, value in copied.items():
            if not isinstance(key, str) or not key.strip():
                raise ValueError("Идентификатор поля должен быть непустой строкой.")
            if value is not None and not isinstance(
                value, str | int | float | bool | Decimal | date | datetime
            ):
                raise ValueError(f"Неподдерживаемый тип значения поля {key}.")
            if isinstance(value, float) and not math.isfinite(value):
                raise ValueError(f"Поле {key} содержит NaN или бесконечность.")
            if isinstance(value, Decimal) and not value.is_finite():
                raise ValueError(f"Поле {key} содержит недопустимое денежное значение.")
        object.__setattr__(self, "article", self.article.strip())
        object.__setattr__(self, "values", MappingProxyType(copied))
```

## src/kuchenland_importer/domain/report.py

```python
"""Import counters represent saved results, never merely planned changes."""

from dataclasses import dataclass
from enum import StrEnum


class Severity(StrEnum):
    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True, slots=True)
class ImportIssue:
    code: str
    message: str
    severity: Severity
    source: str

    def __post_init__(self) -> None:
        if not self.code.strip() or not self.message.strip() or not self.source.strip():
            raise ValueError("Ошибка должна иметь код, сообщение и источник.")


@dataclass(frozen=True, slots=True)
class ImportReport:
    run_id: str
    mails_processed: int = 0
    attachments_processed: int = 0
    products_added: int = 0
    products_updated: int = 0
    products_unchanged: int = 0
    products_skipped: int = 0
    photos_added: int = 0
    photos_replaced: int = 0
    issues: tuple[ImportIssue, ...] = ()

    def __post_init__(self) -> None:
        if not self.run_id.strip():
            raise ValueError("Отчёт должен иметь идентификатор запуска.")
        counts = (
            self.mails_processed,
            self.attachments_processed,
            self.products_added,
            self.products_updated,
            self.products_unchanged,
            self.products_skipped,
            self.photos_added,
            self.photos_replaced,
        )
        if any(type(value) is not int or value < 0 for value in counts):
            raise ValueError("Счётчики должны быть целыми неотрицательными числами.")

    @property
    def error_count(self) -> int:
        return sum(issue.severity == Severity.ERROR for issue in self.issues)

    @property
    def warning_count(self) -> int:
        return sum(issue.severity == Severity.WARNING for issue in self.issues)
```

## src/kuchenland_importer/infrastructure/__init__.py

```python
"""Adapters for configuration, files, logs, and later Office integration."""
```

## src/kuchenland_importer/infrastructure/config_loader.py

```python
"""Strict TOML loading with explicit errors instead of silently ignored typos."""

import tomllib
from collections.abc import Mapping
from pathlib import Path
from typing import cast

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.infrastructure.settings import SeasonRoute, Settings


def _check_keys(data: Mapping[str, object], expected: set[str], section: str) -> None:
    missing = expected - data.keys()
    unknown = data.keys() - expected
    if missing or unknown:
        raise ConfigurationError(
            f"Раздел {section}: отсутствуют {sorted(missing)}, неизвестны {sorted(unknown)}."
        )


def _table(value: object, name: str) -> Mapping[str, object]:
    if not isinstance(value, dict):
        raise ConfigurationError(f"Раздел {name} должен быть таблицей TOML.")
    return cast(Mapping[str, object], value)


def _text(value: object, name: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str) or (not allow_empty and not value.strip()):
        raise ConfigurationError(f"Поле {name} должно быть строкой.")
    return value


def _integer(value: object, name: str) -> int:
    if type(value) is not int:
        raise ConfigurationError(f"Поле {name} должно быть целым числом.")
    return value


def _strings(value: object, name: str) -> tuple[str, ...]:
    if not isinstance(value, list):
        raise ConfigurationError(f"Поле {name} должно быть списком строк.")
    return tuple(_text(item, name) for item in value)


def _path(value: str, base: Path) -> Path:
    path = Path(value).expanduser()
    return (path if path.is_absolute() else base / path).resolve()


def load_settings(
    config_path: Path, *, workbook: Path | None = None, data_dir: Path | None = None
) -> Settings:
    config_path = config_path.resolve()
    try:
        with config_path.open("rb") as stream:
            data: Mapping[str, object] = tomllib.load(stream)
    except (OSError, tomllib.TOMLDecodeError) as error:
        raise ConfigurationError(
            f"Не удалось прочитать настройки {config_path}: {error}"
        ) from error
    _check_keys(data, {"schema_version", "application", "logging", "excel", "routes"}, "root")
    if _integer(data["schema_version"], "schema_version") != 1:
        raise ConfigurationError("Поддерживается schema_version = 1.")
    app = _table(data["application"], "application")
    log = _table(data["logging"], "logging")
    excel = _table(data["excel"], "excel")
    _check_keys(app, {"workbook", "data_dir"}, "application")
    _check_keys(log, {"level", "rotation_mb", "retention_days"}, "logging")
    _check_keys(excel, {"input_extensions", "excluded_sheets"}, "excel")
    raw_routes = data["routes"]
    if not isinstance(raw_routes, list):
        raise ConfigurationError("routes должен быть массивом таблиц TOML.")
    routes: list[SeasonRoute] = []
    for index, raw in enumerate(raw_routes):
        route = _table(raw, f"routes[{index}]")
        _check_keys(route, {"id", "aliases", "calculation_sheet", "photo_sheet"}, "route")
        routes.append(
            SeasonRoute(
                id=_text(route["id"], "route.id"),
                aliases=_strings(route["aliases"], "route.aliases"),
                calculation_sheet=_text(route["calculation_sheet"], "route.calculation_sheet"),
                photo_sheet=_text(route["photo_sheet"], "route.photo_sheet"),
            )
        )
    base = config_path.parent
    workbook_text = _text(app["workbook"], "application.workbook", allow_empty=True)
    configured_workbook = _path(workbook_text, base) if workbook_text.strip() else None
    return Settings(
        workbook=workbook.resolve() if workbook is not None else configured_workbook,
        data_dir=data_dir.resolve()
        if data_dir is not None
        else _path(_text(app["data_dir"], "application.data_dir"), base),
        log_level=_text(log["level"], "logging.level"),
        log_rotation_mb=_integer(log["rotation_mb"], "logging.rotation_mb"),
        log_retention_days=_integer(log["retention_days"], "logging.retention_days"),
        routes=tuple(routes),
        excluded_sheets=_strings(excel["excluded_sheets"], "excel.excluded_sheets"),
        input_extensions=_strings(excel["input_extensions"], "excel.input_extensions"),
    )
```

## src/kuchenland_importer/infrastructure/logging.py

```python
"""Owned Loguru sink with rotation and no captured local-variable values."""

from dataclasses import dataclass
from pathlib import Path

from loguru import logger

from kuchenland_importer.domain.errors import StartupError
from kuchenland_importer.infrastructure.settings import Settings


@dataclass(slots=True)
class LoggingSession:
    sink_id: int | None

    def close(self) -> None:
        if self.sink_id is not None:
            logger.remove(self.sink_id)
            self.sink_id = None


def configure_logging(directory: Path, settings: Settings) -> LoggingSession:
    try:
        sink_id = logger.add(
            directory / "application.log",
            level=settings.log_level,
            rotation=f"{settings.log_rotation_mb} MB",
            retention=f"{settings.log_retention_days} days",
            encoding="utf-8",
            enqueue=False,
            backtrace=False,
            diagnose=False,
            catch=False,
            format=(
                "{time:YYYY-MM-DD HH:mm:ss.SSS ZZ} | {level} | {name}:{function}:{line} | {message}"
            ),
        )
    except (OSError, ValueError) as error:
        raise StartupError(f"Не удалось настроить журнал: {error}") from error
    return LoggingSession(sink_id)
```

## src/kuchenland_importer/infrastructure/paths.py

```python
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
```

## src/kuchenland_importer/infrastructure/settings.py

```python
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
```

## src/kuchenland_importer/presentation/__init__.py

```python
"""User-facing entry points."""
```

## src/kuchenland_importer/presentation/cli.py

```python
"""Stage-one CLI; no unimplemented import command is exposed."""

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from loguru import logger

from kuchenland_importer import __version__
from kuchenland_importer.app import Application
from kuchenland_importer.domain.errors import ApplicationError


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Kuchenland: проверка конфигурации приложения")
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument("--config", type=Path, required=True, help="Путь к настройкам TOML")
    parser.add_argument("--workbook", type=Path, help="Переопределить путь общей книги")
    parser.add_argument("--data-dir", type=Path, help="Переопределить рабочий каталог")
    args = parser.parse_args(argv)
    # The executable owns its logger. Disable the default diagnostic stderr sink.
    logger.remove()
    app: Application | None = None
    try:
        if sys.version_info[:2] != (3, 12):
            print("Требуется Python 3.12.", file=sys.stderr)
            return 2
        app = Application.create(args.config, workbook=args.workbook, data_dir=args.data_dir)
        result = app.diagnose()
        print(f"Kuchenland Importer {__version__} — этап 1")
        for label, messages in (
            ("OK", result.checks),
            ("ПРЕДУПРЕЖДЕНИЕ", result.warnings),
            ("ОШИБКА", result.errors),
        ):
            for message in messages:
                print(f"{label}: {message}")
        return 1 if result.errors else 0
    except ApplicationError as error:
        print(f"Ошибка запуска: {error}", file=sys.stderr)
        return 2
    except Exception:
        if app is not None:
            logger.exception("Непредвиденная ошибка приложения")
            message = f"Непредвиденная ошибка. Журнал: {app.paths.logs / 'application.log'}"
        else:
            message = "Непредвиденная ошибка до открытия журнала. Проверьте настройки и доступы."
        print(message, file=sys.stderr)
        return 3
    finally:
        if app is not None:
            app.close()
```

## tests/integration/test_startup.py

```python
from hashlib import sha256
from importlib.metadata import PackageNotFoundError
from pathlib import Path

import pytest
from loguru import logger

import kuchenland_importer.app as app_module
from kuchenland_importer.app import Application
from kuchenland_importer.presentation.cli import main

DEFAULT = Path(__file__).resolve().parents[2] / "config" / "default.toml"


def test_diagnostics_leave_input_unchanged_and_release_log(tmp_path: Path) -> None:
    # Diagnosis checks accessibility, not Excel validity. This is deliberately not a workbook.
    book = tmp_path / "input.xlsx"
    book.write_bytes(b"immutable source")
    before = sha256(book.read_bytes()).digest()
    app = Application.create(DEFAULT, workbook=book, data_dir=tmp_path / "runtime")
    try:
        result = app.diagnose()
        assert not result.errors
        assert result.warnings
    finally:
        app.close()
    assert sha256(book.read_bytes()).digest() == before
    log = app.paths.logs / "application.log"
    assert "Приложение завершено" in log.read_text(encoding="utf-8")
    log.rename(log.with_suffix(".closed"))


def test_cli_bad_configuration_and_missing_workbook(tmp_path: Path) -> None:
    assert main(["--config", str(tmp_path / "missing.toml")]) == 2
    assert (
        main(
            [
                "--config",
                str(DEFAULT),
                "--data-dir",
                str(tmp_path / "runtime"),
                "--workbook",
                str(tmp_path / "missing.xlsx"),
            ]
        )
        == 1
    )


def test_repeated_startup_has_no_duplicate_file_sink(tmp_path: Path) -> None:
    for _ in range(2):
        app = Application.create(DEFAULT, data_dir=tmp_path / "runtime")
        app.close()
    # Closed application sinks must not receive unrelated messages.
    log = tmp_path / "runtime" / "logs" / "application.log"
    original = log.read_bytes()
    logger.info("outside-session")
    assert log.read_bytes() == original


def test_runtime_path_occupied_by_file(tmp_path: Path) -> None:
    root = tmp_path / "runtime"
    root.write_text("not a directory", encoding="utf-8")
    assert main(["--config", str(DEFAULT), "--data-dir", str(root)]) == 2


def test_missing_dependencies_are_reported(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    def missing_version(package: str) -> str:
        raise PackageNotFoundError(package)

    monkeypatch.setattr(app_module, "version", missing_version)
    app = Application.create(DEFAULT, data_dir=tmp_path / "runtime")
    try:
        result = app.diagnose()
        assert len(result.errors) == 6
        assert all("Не установлена зависимость" in error for error in result.errors)
    finally:
        app.close()
```

## tests/unit/test_models.py

```python
from datetime import UTC, datetime
from decimal import Decimal
from pathlib import Path
from typing import cast

import pytest

from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.product import CellValue, ProductRecord
from kuchenland_importer.domain.report import ImportIssue, ImportReport, Severity


def attachment(tmp_path: Path) -> SavedAttachment:
    mail = MailMetadata("entry", "store", "Тема", datetime(2026, 10, 9, tzinfo=UTC), "Закупщик", "")
    return SavedAttachment(mail, 1, "товары.xlsx", tmp_path / "товары.xlsx", "a" * 64)


def test_article_zeros_and_defensive_copy(tmp_path: Path) -> None:
    values: dict[str, CellValue] = {"price": Decimal("15.20")}
    product = ProductRecord(
        " 001-Ab ", "Поставщик", "Весна 2027", attachment(tmp_path), "Товары", 2, values
    )
    values["price"] = Decimal("999")
    assert product.article == "001-Ab"
    assert product.values["price"] == Decimal("15.20")
    with pytest.raises(TypeError):
        cast(dict[str, CellValue], product.values)["price"] = 0


@pytest.mark.parametrize("value", [float("nan"), float("inf"), Decimal("NaN")])
def test_non_finite_values_rejected(tmp_path: Path, value: CellValue) -> None:
    with pytest.raises(ValueError):
        ProductRecord(
            "001", "Поставщик", "Весна", attachment(tmp_path), "Лист", 2, {"price": value}
        )


def test_received_date_requires_timezone() -> None:
    with pytest.raises(ValueError, match="часовой пояс"):
        MailMetadata("entry", "store", "", datetime(2026, 10, 9), "", "")


def test_report_separates_errors_and_warnings() -> None:
    report = ImportReport(
        "run",
        issues=(
            ImportIssue("NO_PHOTO", "Фото отсутствует", Severity.WARNING, "book.xlsx:2"),
            ImportIssue("NO_SKU", "Нет артикула", Severity.ERROR, "book.xlsx:3"),
        ),
    )
    assert report.error_count == 1
    assert report.warning_count == 1


def test_report_rejects_negative_count() -> None:
    with pytest.raises(ValueError):
        ImportReport("run", products_added=-1)
```

## tests/unit/test_settings.py

```python
from pathlib import Path

import pytest

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.settings import normalize_alias

DEFAULT = Path(__file__).resolve().parents[2] / "config" / "default.toml"


def test_default_routes_and_excluded_sheet() -> None:
    settings = load_settings(DEFAULT)
    assert [route.id for route in settings.routes] == [
        "spring",
        "easter",
        "summer",
        "autumn",
        "winter",
    ]
    assert settings.excluded_sheets == ("удалено",)
    assert settings.workbook is None
    assert settings.data_dir == DEFAULT.parent.parent / ".runtime"


@pytest.mark.parametrize("text", [" Весна – лето ", "ВЕСНА-ЛЕТО", "Весна‑лето"])
def test_alias_normalization(text: str) -> None:
    assert normalize_alias(text) == "весна-лето"


@pytest.mark.parametrize(
    ("before", "after", "message"),
    [
        ('aliases = ["Пасха"]', 'aliases = ["ВЕСНА"]', "неоднозначно"),
        ('id = "winter"', 'id = "spring"', "повторно"),
        ('photo_sheet = "ФОТО зима"', 'photo_sheet = "ФОТО весна"', "неоднозначно"),
        ('calculation_sheet = "Зима"', 'calculation_sheet = "удалено"', "неоднозначно"),
        ('calculation_sheet = "Зима"', 'calculation_sheet = "Зима/2027"', "имя вкладки"),
        ("rotation_mb = 10", "rotation_mb = true", "целым числом"),
        ("rotation_mb = 10", "rotation_mb = 0", "положительными"),
        ('level = "INFO"', 'level = "INF"', "уровень"),
        ("schema_version = 1", "schema_version = 2", "schema_version"),
        ("retention_days = 30", "retentoin_days = 30", "неизвестны"),
        ('".xlsb"', '".csv"', "форматы"),
    ],
)
def test_invalid_configuration_is_rejected(
    tmp_path: Path, before: str, after: str, message: str
) -> None:
    config = tmp_path / "settings.toml"
    config.write_text(DEFAULT.read_text(encoding="utf-8").replace(before, after), encoding="utf-8")
    with pytest.raises(ConfigurationError, match=message):
        load_settings(config)


def test_override_paths(tmp_path: Path) -> None:
    book = tmp_path / "book.xlsx"
    settings = load_settings(DEFAULT, workbook=book, data_dir=tmp_path / "data")
    assert settings.workbook == book
    assert settings.data_dir == tmp_path / "data"


def test_custom_season_without_code_changes(tmp_path: Path) -> None:
    config = tmp_path / "settings.toml"
    config.write_text(
        DEFAULT.read_text(encoding="utf-8")
        + """
[[routes]]
id = "new_year"
aliases = ["Новый год"]
calculation_sheet = "Новый год"
photo_sheet = "ФОТО Новый год"
""",
        encoding="utf-8",
    )
    assert load_settings(config).routes[-1].id == "new_year"


def test_missing_and_malformed_config(tmp_path: Path) -> None:
    config = tmp_path / "bad.toml"
    with pytest.raises(ConfigurationError, match="прочитать"):
        load_settings(config)
    config.write_text("broken = [", encoding="utf-8")
    with pytest.raises(ConfigurationError, match="прочитать"):
        load_settings(config)
```

## requirements-lock.txt

```text
# Tested development environment: Windows, Python 3.12.8, stage 1, 2026-10-09.
# Includes runtime and QA dependencies; Office integration is not yet tested.
colorama==0.4.6
et_xmlfile==2.0.0
iniconfig==2.3.1
loguru==0.7.3
mypy==1.15.0
mypy_extensions==1.1.0
numpy==2.2.6
openpyxl==3.1.5
packaging==26.3
pandas==2.2.3
pillow==11.2.1
pluggy==1.6.0
pytest==8.3.5
python-dateutil==2.9.0.post0
pytz==2025.2
pywin32==312
ruff==0.11.13
six==1.17.0
typing_extensions==4.16.0
tzdata==2025.2
win32_setctime==1.2.0
xlwings==0.37.5
```
