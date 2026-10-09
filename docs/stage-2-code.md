# Полные файлы этапа 2

Версия 0.2.0. Полные актуальные тексты для ревью, включая изменённые файлы этапа 1. Исторический снимок этапа 1 сохранён отдельно.

## main.py

````python
"""Launch the application from a source checkout without installation."""

import sys
from pathlib import Path

if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
    from kuchenland_importer.presentation.cli import main

    raise SystemExit(main())
````

## pyproject.toml

````toml
[build-system]
requires = ["setuptools>=75,<83"]
build-backend = "setuptools.build_meta"

[project]
name = "kuchenland-importer"
version = "0.2.0"
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

[[tool.mypy.overrides]]
module = ["pythoncom", "pywintypes", "win32com.*"]
ignore_missing_imports = true
````

## requirements.txt

````text
# Runtime dependencies; tested exact versions are recorded in requirements-lock.txt.
pywin32>=306; sys_platform == "win32"
xlwings>=0.33,<1
openpyxl>=3.1.5,<4
pandas>=2.2,<4
Pillow>=10.4,<13
loguru>=0.7.2,<1
# tkinter ships with the standard Windows Python distribution.
````

## requirements-dev.txt

````text
-r requirements.txt
pytest>=8.3,<10
ruff>=0.6,<1
mypy>=1.11,<2
````

## requirements-lock.txt

````text
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
````

## config/default.toml

````toml
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
````

## src/kuchenland_importer/__init__.py

````python
"""Kuchenland procurement importer."""

__version__ = "0.2.0"
````

## src/kuchenland_importer/app.py

````python
"""Composition root for diagnostics and Outlook source capture."""

from dataclasses import dataclass
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Self

from loguru import logger

from kuchenland_importer.application.capture_mail import CaptureMail, CaptureOutcome
from kuchenland_importer.infrastructure.capture_manifest import JsonCaptureManifestStore
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
        warnings.append("Диагностика не подключается к Office; импорт товаров ещё не реализован.")
        logger.info("Диагностика: ошибок {}, предупреждений {}", len(errors), len(warnings))
        return DiagnosticResult(tuple(checks), tuple(warnings), tuple(errors))

    def capture_outlook(self) -> CaptureOutcome:
        from kuchenland_importer.infrastructure.outlook.source import OutlookMailSource

        use_case = CaptureMail(OutlookMailSource(), JsonCaptureManifestStore())
        return use_case.execute(self.paths.work, self.settings.input_extensions)

    def close(self) -> None:
        logger.info("Приложение завершено.")
        self.logging.close()
````

## src/kuchenland_importer/application/__init__.py

````python
"""Office-independent use cases and ports."""
````

## src/kuchenland_importer/application/capture_mail.py

````python
"""Collect source files in an isolated run without opening the destination book."""

from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from loguru import logger

from kuchenland_importer.application.ports import CaptureManifestStore, MailSource
from kuchenland_importer.domain.capture import MailCapture
from kuchenland_importer.domain.errors import CaptureStorageError


@dataclass(frozen=True, slots=True)
class CaptureOutcome:
    capture: MailCapture
    manifest: Path


@dataclass(slots=True)
class CaptureMail:
    source: MailSource
    manifests: CaptureManifestStore

    def execute(self, work_dir: Path, extensions: tuple[str, ...]) -> CaptureOutcome:
        run_id = uuid4().hex
        directory = work_dir / run_id
        try:
            directory.mkdir(exist_ok=False)
        except OSError as error:
            raise CaptureStorageError("Не удалось создать каталог запуска.") from error
        logger.info("Получение писем: запуск {}", run_id)
        capture = self.source.capture(run_id, directory, extensions)
        for issue in capture.issues:
            logger.warning("Запуск {}: {} ({})", run_id, issue.code, issue.source)
        manifest = self.manifests.save(capture)
        logger.info(
            "Запуск {}: писем {}, вложений {}, ошибок {}",
            run_id,
            len(capture.mails),
            len(capture.attachments),
            capture.error_count,
        )
        return CaptureOutcome(capture, manifest)
````

## src/kuchenland_importer/application/ports.py

````python
"""Ports for collecting mail and recording its provenance."""

from pathlib import Path
from typing import Protocol

from kuchenland_importer.domain.capture import MailCapture


class MailSource(Protocol):
    def capture(self, run_id: str, directory: Path, extensions: tuple[str, ...]) -> MailCapture: ...


class CaptureManifestStore(Protocol):
    def save(self, capture: MailCapture) -> Path: ...
````

## src/kuchenland_importer/domain/__init__.py

````python
"""Office-independent business models and invariants."""
````

## src/kuchenland_importer/domain/capture.py

````python
"""Result of collecting source mail; no Excel rows have been imported yet."""

from dataclasses import dataclass
from pathlib import Path

from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.report import ImportIssue, Severity


@dataclass(frozen=True, slots=True)
class MailCapture:
    run_id: str
    directory: Path
    selected_count: int
    mails: tuple[MailMetadata, ...]
    attachments: tuple[SavedAttachment, ...]
    issues: tuple[ImportIssue, ...]
    selection_positions: tuple[int, ...]

    def __post_init__(self) -> None:
        if not self.run_id.strip() or not self.directory.is_absolute():
            raise ValueError("Запуск должен иметь идентификатор и абсолютный каталог.")
        if type(self.selected_count) is not int or self.selected_count < len(self.mails):
            raise ValueError("Число элементов выделения не может быть меньше числа писем.")
        identities = {(mail.store_id, mail.entry_id) for mail in self.mails}
        if len(identities) != len(self.mails):
            raise ValueError("В результате получения писем обнаружены дубли.")
        if len(self.selection_positions) != len(self.mails) or any(
            index < 1 or index > self.selected_count for index in self.selection_positions
        ):
            raise ValueError("Позиции писем должны соответствовать исходному выделению.")
        for attachment in self.attachments:
            if (attachment.mail.store_id, attachment.mail.entry_id) not in identities:
                raise ValueError("Вложение не связано с полученным письмом.")
            if not attachment.path.is_relative_to(self.directory):
                raise ValueError("Вложение находится вне каталога запуска.")

    @property
    def error_count(self) -> int:
        return sum(issue.severity == Severity.ERROR for issue in self.issues)
````

## src/kuchenland_importer/domain/errors.py

````python
"""Errors that can be translated into actionable user messages."""


class ApplicationError(Exception):
    """Base class for expected application failures."""


class ConfigurationError(ApplicationError):
    """Configuration is missing, inconsistent, or malformed."""


class StartupError(ApplicationError):
    """Application resources cannot be initialized."""


class OutlookError(ApplicationError):
    """Classic Outlook is unavailable or its selection cannot be obtained."""


class AttachmentError(ApplicationError):
    """An attachment could not be safely saved and verified."""


class CaptureStorageError(ApplicationError):
    """Source capture directories or manifest could not be written."""
````

## src/kuchenland_importer/domain/mail.py

````python
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
````

## src/kuchenland_importer/domain/product.py

````python
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
````

## src/kuchenland_importer/domain/report.py

````python
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
````

## src/kuchenland_importer/infrastructure/__init__.py

````python
"""Adapters for configuration, files, logs, and later Office integration."""
````

## src/kuchenland_importer/infrastructure/capture_manifest.py

````python
"""UTF-8 source manifest; mail content is never included in diagnostic logs."""

import json
from pathlib import Path

from kuchenland_importer.domain.capture import MailCapture
from kuchenland_importer.domain.errors import CaptureStorageError


class JsonCaptureManifestStore:
    def save(self, capture: MailCapture) -> Path:
        final = capture.directory / "manifest.json"
        partial = capture.directory / "manifest.json.part"
        mail_indices = {(mail.store_id, mail.entry_id): i for i, mail in enumerate(capture.mails)}
        data = {
            "schema_version": 1,
            "stage": "mail_capture",
            "run_id": capture.run_id,
            "selected_count": capture.selected_count,
            "mails": [
                {
                    "selection_index": capture.selection_positions[index],
                    "entry_id": mail.entry_id,
                    "store_id": mail.store_id,
                    "subject": mail.subject,
                    "received_at": mail.received_at.isoformat(),
                    "sender_name": mail.sender_name,
                    "sender_address": mail.sender_address,
                }
                for index, mail in enumerate(capture.mails)
            ],
            "attachments": [
                {
                    "mail_index": mail_indices[(item.mail.store_id, item.mail.entry_id)],
                    "attachment_index": item.attachment_index,
                    "original_name": item.original_name,
                    "path": item.path.relative_to(capture.directory).as_posix(),
                    "sha256": item.sha256,
                }
                for item in capture.attachments
            ],
            "issues": [
                {
                    "code": issue.code,
                    "message": issue.message,
                    "severity": issue.severity.value,
                    "source": issue.source,
                }
                for issue in capture.issues
            ],
        }
        try:
            with partial.open("x", encoding="utf-8") as stream:
                json.dump(data, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
            partial.rename(final)
        except OSError as error:
            raise CaptureStorageError(
                f"Отчёт не сохранён. Полученные файлы оставлены в {capture.directory}."
            ) from error
        return final
````

## src/kuchenland_importer/infrastructure/config_loader.py

````python
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
````

## src/kuchenland_importer/infrastructure/logging.py

````python
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
````

## src/kuchenland_importer/infrastructure/outlook/__init__.py

````python
"""Classic Outlook COM adapter. No email mutation or sending operations."""
````

## src/kuchenland_importer/infrastructure/outlook/attachment_writer.py

````python
"""Save under generated names, hash actual bytes, publish only complete files."""

from hashlib import file_digest
from pathlib import Path, PureWindowsPath

import pywintypes

from kuchenland_importer.domain.errors import AttachmentError
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.infrastructure.outlook.com_types import ComAttachment


def attachment_extension(name: str) -> str:
    return PureWindowsPath(name).suffix.casefold()


def save_attachment(
    attachment: ComAttachment, mail: MailMetadata, index: int, directory: Path, name: str
) -> SavedAttachment:
    suffix = attachment_extension(name)
    final = directory / f"attachment-{index:04d}{suffix}"
    partial = directory / f"attachment-{index:04d}{suffix}.part"
    if final.exists() or partial.exists():
        raise AttachmentError("Файл с таким индексом уже существует; перезапись запрещена.")
    try:
        attachment.SaveAsFile(str(partial))
        if partial.stat().st_size == 0:
            raise AttachmentError("Outlook сохранил пустое вложение.")
        with partial.open("rb") as stream:
            digest = file_digest(stream, "sha256").hexdigest()
        partial.rename(final)
    except (OSError, pywintypes.com_error, AttachmentError) as error:
        try:
            partial.unlink(missing_ok=True)
        except OSError as cleanup_error:
            raise AttachmentError(
                "Вложение не сохранено; временный файл не удалось удалить."
            ) from cleanup_error
        raise AttachmentError("Не удалось сохранить и проверить вложение.") from error
    return SavedAttachment(mail, index, name, final, digest)
````

## src/kuchenland_importer/infrastructure/outlook/com_types.py

````python
"""Small typed boundary around Outlook's dynamically dispatched objects."""

from typing import Protocol


class PropertyAccessor(Protocol):
    def GetProperty(self, name: str) -> object: ...


class ComAttachment(Protocol):
    @property
    def FileName(self) -> str: ...

    def SaveAsFile(self, path: str) -> None: ...


class Attachments(Protocol):
    @property
    def Count(self) -> int: ...

    def Item(self, index: int) -> ComAttachment: ...


class Folder(Protocol):
    @property
    def StoreID(self) -> str: ...


class OutlookItem(Protocol):
    @property
    def Class(self) -> int: ...


class MailItem(OutlookItem, Protocol):
    @property
    def EntryID(self) -> str: ...

    @property
    def Parent(self) -> Folder: ...

    @property
    def Subject(self) -> str: ...

    @property
    def SenderName(self) -> str: ...

    @property
    def SenderEmailAddress(self) -> str: ...

    @property
    def PropertyAccessor(self) -> PropertyAccessor: ...

    @property
    def Attachments(self) -> Attachments: ...


class Selection(Protocol):
    @property
    def Count(self) -> int: ...

    def Item(self, index: int) -> object: ...


class Explorer(Protocol):
    @property
    def Selection(self) -> Selection: ...


class OutlookApplication(Protocol):
    def ActiveExplorer(self) -> Explorer | None: ...
````

## src/kuchenland_importer/infrastructure/outlook/metadata.py

````python
"""Read identity and delivery time; the MAPI PT_SYSTIME value is UTC."""

from datetime import UTC, datetime

from kuchenland_importer.domain.mail import MailMetadata
from kuchenland_importer.infrastructure.outlook.com_types import MailItem

DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"


def read_metadata(item: MailItem) -> MailMetadata:
    received = item.PropertyAccessor.GetProperty(DELIVERY_TIME)
    if not isinstance(received, datetime):
        raise ValueError("Outlook не предоставил дату получения письма.")
    # PropertyAccessor returns UTC wall-clock values, unlike the local ReceivedTime property.
    received_utc = received.replace(tzinfo=UTC)
    return MailMetadata(
        entry_id=item.EntryID,
        store_id=item.Parent.StoreID,
        subject=item.Subject or "",
        received_at=received_utc,
        sender_name=item.SenderName or "",
        sender_address=item.SenderEmailAddress or "",
    )
````

## src/kuchenland_importer/infrastructure/outlook/session.py

````python
"""Initialize COM in the calling thread and attach only to a running Outlook."""

import sys
from collections.abc import Iterator
from contextlib import contextmanager
from typing import cast

import pythoncom
import pywintypes
import win32com.client

from kuchenland_importer.domain.errors import OutlookError
from kuchenland_importer.infrastructure.outlook.com_types import OutlookApplication


@contextmanager
def outlook_session() -> Iterator[OutlookApplication]:
    if sys.platform != "win32":
        raise OutlookError("Подключение к Outlook поддерживается только в Windows.")
    app: OutlookApplication | None = None
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except pywintypes.com_error as error:
        raise OutlookError("Не удалось инициализировать COM в текущем потоке.") from error
    try:
        try:
            app = cast(OutlookApplication, win32com.client.GetActiveObject("Outlook.Application"))
        except pywintypes.com_error as error:
            raise OutlookError(
                "Классический Outlook недоступен. Откройте Outlook с настроенной учётной "
                "записью и выделите письма. Новый Outlook этим способом не поддерживается."
            ) from error
        yield app
    finally:
        app = None
        pythoncom.CoUninitialize()
````

## src/kuchenland_importer/infrastructure/outlook/source.py

````python
"""Capture the Explorer selection once; isolate errors to items and attachments."""

from pathlib import Path
from typing import cast

import pywintypes

from kuchenland_importer.domain.capture import MailCapture
from kuchenland_importer.domain.errors import AttachmentError, OutlookError
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.infrastructure.outlook.attachment_writer import (
    attachment_extension,
    save_attachment,
)
from kuchenland_importer.infrastructure.outlook.com_types import (
    MailItem,
    OutlookApplication,
    OutlookItem,
    Selection,
)
from kuchenland_importer.infrastructure.outlook.metadata import read_metadata
from kuchenland_importer.infrastructure.outlook.session import outlook_session

OL_MAIL = 43


def _snapshot(selection: Selection) -> tuple[int, list[tuple[int, object]], list[ImportIssue]]:
    count = selection.Count
    if count == 0:
        raise OutlookError("В Outlook не выделены письма.")
    items: list[tuple[int, object]] = []
    issues: list[ImportIssue] = []
    for index in range(1, count + 1):
        try:
            items.append((index, selection.Item(index)))
        except pywintypes.com_error:
            issues.append(
                ImportIssue(
                    "SELECTION_ITEM_FAILED",
                    "Элемент выделения недоступен.",
                    Severity.ERROR,
                    f"selection:{index}",
                )
            )
    return count, items, issues


def capture_selection(
    selection: Selection, run_id: str, directory: Path, extensions: tuple[str, ...]
) -> MailCapture:
    count, items, issues = _snapshot(selection)
    mails: list[MailMetadata] = []
    positions: list[int] = []
    attachments: list[SavedAttachment] = []
    identities: set[tuple[str, str]] = set()
    for ordinal, item in items:
        source = f"selection:{ordinal}"
        try:
            if cast(OutlookItem, item).Class != OL_MAIL:
                issues.append(
                    ImportIssue(
                        "NOT_MAIL",
                        "Выделенный элемент не является письмом.",
                        Severity.WARNING,
                        source,
                    )
                )
                continue
            mail_item = cast(MailItem, item)
            mail = read_metadata(mail_item)
            identity = (mail.store_id, mail.entry_id)
            if identity in identities:
                issues.append(
                    ImportIssue(
                        "DUPLICATE_MAIL", "Повторное письмо пропущено.", Severity.WARNING, source
                    )
                )
                continue
            identities.add(identity)
            mails.append(mail)
            positions.append(ordinal)
            mail_dir = directory / f"mail-{ordinal:04d}"
            mail_dir.mkdir(exist_ok=False)
            collection = mail_item.Attachments
            attachment_count = collection.Count
        except (pywintypes.com_error, OSError, ValueError, AttributeError):
            issues.append(
                ImportIssue(
                    "MAIL_READ_FAILED", "Не удалось прочитать письмо.", Severity.ERROR, source
                )
            )
            continue
        found = 0
        for index in range(1, attachment_count + 1):
            try:
                attachment = collection.Item(index)
                name = attachment.FileName
                if attachment_extension(name) not in extensions:
                    continue
                found += 1
                attachments.append(save_attachment(attachment, mail, index, mail_dir, name))
            except (pywintypes.com_error, OSError, ValueError, AttachmentError):
                issues.append(
                    ImportIssue(
                        "ATTACHMENT_SAVE_FAILED",
                        "Excel-вложение не сохранено.",
                        Severity.ERROR,
                        f"{source}/attachment:{index}",
                    )
                )
        if found == 0:
            issues.append(
                ImportIssue(
                    "NO_EXCEL", "У письма нет доступных Excel-вложений.", Severity.WARNING, source
                )
            )
    return MailCapture(
        run_id, directory, count, tuple(mails), tuple(attachments), tuple(issues), tuple(positions)
    )


class OutlookMailSource:
    def capture(self, run_id: str, directory: Path, extensions: tuple[str, ...]) -> MailCapture:
        with outlook_session() as app:
            try:
                return self._capture_active_explorer(app, run_id, directory, extensions)
            finally:
                del app

    @staticmethod
    def _capture_active_explorer(
        app: OutlookApplication, run_id: str, directory: Path, extensions: tuple[str, ...]
    ) -> MailCapture:
        try:
            explorer = app.ActiveExplorer()
            if explorer is None:
                raise OutlookError("Откройте основное окно Outlook и выделите письма в списке.")
            selection = explorer.Selection
        except pywintypes.com_error as error:
            raise OutlookError("Не удалось прочитать выделение в основном окне Outlook.") from error
        try:
            return capture_selection(selection, run_id, directory, extensions)
        except pywintypes.com_error as error:
            raise OutlookError("Выделение Outlook стало недоступно. Повторите запуск.") from error
````

## src/kuchenland_importer/infrastructure/paths.py

````python
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
````

## src/kuchenland_importer/infrastructure/settings.py

````python
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
````

## src/kuchenland_importer/presentation/__init__.py

````python
"""User-facing entry points."""
````

## src/kuchenland_importer/presentation/cli.py

````python
"""Diagnostics and source collection; Excel import is not exposed yet."""

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from loguru import logger

from kuchenland_importer import __version__
from kuchenland_importer.app import Application
from kuchenland_importer.domain.errors import ApplicationError


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Kuchenland: диагностика и получение писем")
    parser.add_argument(
        "command", nargs="?", default="diagnose", choices=("diagnose", "capture-outlook")
    )
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
        if args.command == "capture-outlook":
            outcome = app.capture_outlook()
            capture = outcome.capture
            print(f"Получено писем: {len(capture.mails)} из {capture.selected_count} элементов")
            print(f"Сохранено Excel-вложений: {len(capture.attachments)}")
            for issue in capture.issues:
                print(f"{issue.severity.value}: {issue.source}: {issue.message}")
            print(f"Отчёт: {outcome.manifest}")
            print("Получены исходные данные; общая книга не изменялась.")
            return 1 if capture.error_count or not capture.attachments else 0
        result = app.diagnose()
        print(f"Kuchenland Importer {__version__} — диагностика")
        for label, messages in (
            ("OK", result.checks),
            ("ПРЕДУПРЕЖДЕНИЕ", result.warnings),
            ("ОШИБКА", result.errors),
        ):
            for message in messages:
                print(f"{label}: {message}")
        return 1 if result.errors else 0
    except ApplicationError as error:
        if app is not None:
            logger.warning("Ошибка приложения: {}", error)
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
````

## tests/integration/test_startup.py

````python
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
````

## tests/unit/test_models.py

````python
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
````

## tests/unit/test_outlook_capture.py

````python
from contextlib import contextmanager
from datetime import UTC, datetime, timedelta, timezone
from hashlib import sha256
from pathlib import Path
from types import SimpleNamespace
from typing import cast
from unittest.mock import Mock

import pytest
import pywintypes

from kuchenland_importer.application.capture_mail import CaptureMail
from kuchenland_importer.domain.errors import AttachmentError, CaptureStorageError, OutlookError
from kuchenland_importer.infrastructure.capture_manifest import JsonCaptureManifestStore
from kuchenland_importer.infrastructure.outlook import session, source
from kuchenland_importer.infrastructure.outlook.attachment_writer import save_attachment
from kuchenland_importer.infrastructure.outlook.com_types import (
    MailItem,
    OutlookApplication,
    Selection,
)
from kuchenland_importer.infrastructure.outlook.metadata import read_metadata
from kuchenland_importer.infrastructure.outlook.source import capture_selection


def com_failure() -> pywintypes.com_error:
    return pywintypes.com_error(-2147221021, "test COM failure", None, None)


class FakeAttachment:
    def __init__(self, name: str, content: bytes = b"Excel bytes", fail: bool = False):
        self.FileName = name
        self.content = content
        self.fail = fail

    def SaveAsFile(self, path: str) -> None:
        Path(path).write_bytes(self.content)
        if self.fail:
            raise com_failure()


class FakeCollection:
    def __init__(self, items: list[object]):
        self.items = items
        self.Count = len(items)

    def Item(self, index: int) -> object:
        item = self.items[index - 1]
        if isinstance(item, Exception):
            raise item
        return item


def mail(
    entry: str = "entry", attachments: list[object] | None = None, store: str = "store"
) -> SimpleNamespace:
    return SimpleNamespace(
        Class=43,
        EntryID=entry,
        Parent=SimpleNamespace(StoreID=store),
        Subject="Заказник — весна",
        SenderName="Закупщик",
        SenderEmailAddress="sender@example.test",
        PropertyAccessor=SimpleNamespace(GetProperty=lambda _: datetime(2026, 10, 9, 10)),
        Attachments=FakeCollection(attachments or []),
    )


def collect(tmp_path: Path, items: list[object]):
    return capture_selection(cast(Selection, FakeCollection(items)), "run", tmp_path, (".xlsx",))


def test_mixed_selection_and_untrusted_names(tmp_path: Path) -> None:
    capture = collect(
        tmp_path,
        [
            SimpleNamespace(Class=26),
            mail(attachments=[FakeAttachment(r"..\..\CON.XLSX"), FakeAttachment("ignore.pdf")]),
            mail("second", [FakeAttachment(r"..\..\CON.XLSX")]),
        ],
    )
    assert capture.selected_count == 3
    assert len(capture.mails) == 2
    assert capture.selection_positions == (2, 3)
    assert len(capture.attachments) == 2
    paths = [item.path for item in capture.attachments]
    assert paths[0] != paths[1]
    assert all(path.is_relative_to(tmp_path) for path in paths)
    assert all(path.name == "attachment-0001.xlsx" for path in paths)
    assert capture.attachments[0].sha256 == sha256(b"Excel bytes").hexdigest()
    assert capture.attachments[0].original_name == r"..\..\CON.XLSX"
    assert capture.issues[0].code == "NOT_MAIL"


def test_one_failed_attachment_does_not_block_other_files(tmp_path: Path) -> None:
    capture = collect(
        tmp_path,
        [
            mail(
                attachments=[
                    FakeAttachment("bad.xlsx", fail=True),
                    FakeAttachment("empty.xlsx", b""),
                    FakeAttachment("good.xlsx"),
                ]
            )
        ],
    )
    assert len(capture.attachments) == 1
    assert capture.attachments[0].attachment_index == 3
    assert capture.error_count == 2
    assert not list(tmp_path.rglob("*.part"))


def test_duplicate_identity_and_same_id_in_another_store(tmp_path: Path) -> None:
    capture = collect(tmp_path, [mail(), mail(), mail(store="another-store")])
    assert len(capture.mails) == 2
    assert [issue.code for issue in capture.issues].count("DUPLICATE_MAIL") == 1


def test_snapshot_item_failure_and_bad_metadata_continue(tmp_path: Path) -> None:
    broken = mail("bad")
    broken.PropertyAccessor = SimpleNamespace(GetProperty=lambda _: None)
    capture = collect(tmp_path, [com_failure(), broken, mail("good")])
    assert len(capture.mails) == 1
    assert capture.error_count == 2
    assert {issue.code for issue in capture.issues} >= {
        "SELECTION_ITEM_FAILED",
        "MAIL_READ_FAILED",
        "NO_EXCEL",
    }


def test_empty_selection_is_actionable(tmp_path: Path) -> None:
    with pytest.raises(OutlookError, match="не выделены"):
        collect(tmp_path, [])


def test_utc_delivery_time_is_not_converted_twice(tmp_path: Path) -> None:
    item = mail()
    # PT_SYSTIME is UTC wall-clock even if a wrapper adds a tzinfo label.
    item.PropertyAccessor = SimpleNamespace(
        GetProperty=lambda _: datetime(2026, 10, 9, 10, tzinfo=timezone(timedelta(hours=3)))
    )
    assert collect(tmp_path, [item]).mails[0].received_at == datetime(2026, 10, 9, 10, tzinfo=UTC)


def test_com_cleanup_after_connection_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    initialize, uninitialize = Mock(), Mock()
    monkeypatch.setattr(session.pythoncom, "CoInitializeEx", initialize)
    monkeypatch.setattr(session.pythoncom, "CoUninitialize", uninitialize)
    monkeypatch.setattr(session.win32com.client, "GetActiveObject", Mock(side_effect=com_failure()))
    with pytest.raises(OutlookError, match="Классический Outlook"):
        with session.outlook_session():
            pytest.fail("Unavailable Outlook must not yield")
    initialize.assert_called_once()
    uninitialize.assert_called_once()


def test_com_initialization_failure_is_not_uninitialized(monkeypatch: pytest.MonkeyPatch) -> None:
    uninitialize = Mock()
    monkeypatch.setattr(session.pythoncom, "CoInitializeEx", Mock(side_effect=com_failure()))
    monkeypatch.setattr(session.pythoncom, "CoUninitialize", uninitialize)
    with pytest.raises(OutlookError, match="инициализировать COM"):
        with session.outlook_session():
            pytest.fail("Failed initialization must not yield")
    uninitialize.assert_not_called()


def test_com_cleanup_after_body_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    uninitialize = Mock()
    monkeypatch.setattr(session.pythoncom, "CoInitializeEx", Mock())
    monkeypatch.setattr(session.pythoncom, "CoUninitialize", uninitialize)
    monkeypatch.setattr(session.win32com.client, "GetActiveObject", Mock(return_value=object()))
    with pytest.raises(RuntimeError, match="body failure"):
        with session.outlook_session():
            raise RuntimeError("body failure")
    uninitialize.assert_called_once()


def test_attachment_never_overwrites_existing_file(tmp_path: Path) -> None:
    existing = tmp_path / "attachment-0001.xlsx"
    existing.write_bytes(b"original")
    with pytest.raises(AttachmentError, match="перезапись запрещена"):
        save_attachment(
            FakeAttachment("new.xlsx"),
            read_metadata(cast(MailItem, mail())),
            1,
            tmp_path,
            "new.xlsx",
        )
    assert existing.read_bytes() == b"original"


def test_manifest_failure_keeps_saved_attachment(tmp_path: Path) -> None:
    capture = collect(tmp_path, [mail(attachments=[FakeAttachment("good.xlsx")])])
    report = tmp_path / "manifest.json"
    report.write_text("existing report", encoding="utf-8")
    with pytest.raises(CaptureStorageError):
        JsonCaptureManifestStore().save(capture)
    assert report.read_text(encoding="utf-8") == "existing report"
    assert capture.attachments[0].path.read_bytes() == b"Excel bytes"


def test_no_explorer_is_actionable(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    @contextmanager
    def fake_session():
        yield cast(OutlookApplication, SimpleNamespace(ActiveExplorer=lambda: None))

    monkeypatch.setattr(source, "outlook_session", fake_session)
    with pytest.raises(OutlookError, match="основное окно"):
        source.OutlookMailSource().capture("run", tmp_path, (".xlsx",))


def test_use_case_writes_manifest_and_preserves_original_name(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import json

    selection = FakeCollection([mail(attachments=[FakeAttachment("Поставщик.xlsx")])])
    app = SimpleNamespace(ActiveExplorer=lambda: SimpleNamespace(Selection=selection))

    @contextmanager
    def fake_session():
        yield cast(OutlookApplication, app)

    monkeypatch.setattr(source, "outlook_session", fake_session)
    service = CaptureMail(source.OutlookMailSource(), JsonCaptureManifestStore())
    first = service.execute(tmp_path, (".xlsx",))
    second = service.execute(tmp_path, (".xlsx",))
    assert first.capture.directory != second.capture.directory
    manifest = json.loads(first.manifest.read_text(encoding="utf-8"))
    assert manifest["stage"] == "mail_capture"
    assert manifest["mails"][0]["received_at"].endswith("+00:00")
    assert manifest["attachments"][0]["original_name"] == "Поставщик.xlsx"
    assert manifest["attachments"][0]["path"] == "mail-0001/attachment-0001.xlsx"
    assert not list(tmp_path.rglob("*.part"))


def test_cli_capture_does_not_require_destination_workbook(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    from kuchenland_importer.presentation.cli import main

    @contextmanager
    def fake_session():
        selection = FakeCollection([mail(attachments=[FakeAttachment("товары.xlsx")])])
        yield cast(
            OutlookApplication,
            SimpleNamespace(ActiveExplorer=lambda: SimpleNamespace(Selection=selection)),
        )

    monkeypatch.setattr(source, "outlook_session", fake_session)
    config = Path(__file__).resolve().parents[2] / "config" / "default.toml"
    missing_book = tmp_path / "does-not-exist.xlsx"
    assert (
        main(
            [
                "capture-outlook",
                "--config",
                str(config),
                "--data-dir",
                str(tmp_path / "data"),
                "--workbook",
                str(missing_book),
            ]
        )
        == 0
    )
    assert "Сохранено Excel-вложений: 1" in capsys.readouterr().out
    assert not missing_book.exists()
````

## tests/unit/test_settings.py

````python
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
````

## README.md

````markdown
# Kuchenland Importer

Windows-приложение для импорта товаров из выделенных писем классического Outlook
в локальную общую книгу Excel. Python 3.12, Office Professional Plus 2024 / Microsoft 365.

## Текущий статус

Реализованы этапы 1–2: основа приложения и получение выделенных писем классического
Outlook с сохранением Excel-вложений и отчётом. Автоматические тесты проверены;
приёмка на настоящем Outlook отложена до устройства с настроенной учётной записью.
Импорт товаров в Excel, работа с фото, GUI и EXE ещё не реализованы.
Приложение пока не готово к производственному использованию.

## Установка для разработки

В PowerShell из корня проекта:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

`requirements.txt` содержит рабочие зависимости; `requirements-dev.txt` добавляет
проверки качества. Если присутствует `requirements-lock.txt`, он фиксирует версии
проверенного окружения и устанавливается вместо `requirements-dev.txt` для его
точного воспроизведения. `tkinter` входит в стандартную Windows-установку Python.

## Запуск этапа 1

```powershell
.\.venv\Scripts\python.exe main.py --config config/default.toml
```

Команда проверяет настройки, создаёт рабочие каталоги и журнал, проверяет наличие
библиотек. Если общая книга не выбрана, сообщает предупреждение. Не подключается
к Outlook/Excel, не изменяет и не сохраняет книги.

Для проверки своей книги передайте абсолютный путь:

```powershell
.\.venv\Scripts\python.exe main.py --config config/default.toml --workbook "D:\вайбкодин\Почта_сводник\=РАССЧИТАНО= ВЕСНА  ВЕСНА-ЛЕТО и ПАСХА 08.10.2026.xlsx"
```

Для персональных настроек скопируйте `config/default.toml` в `config/local.toml`
и задайте `application.workbook`. `local.toml` исключён из Git.
Каждый конфигурационный файл — полный документ; автоматического слияния нет.
Относительные пути в TOML считаются от папки файла настроек. Относительные пути
аргументов CLI считаются от текущего рабочего каталога.

`--data-dir` переопределяет рабочий каталог. По умолчанию при запуске из исходников
это `.runtime/`, содержащий `logs`, `backup`, `work`, `reports`.
Для будущей установленной версии пользовательский каталог будет определён на этапе 8.

Коды завершения: `0` — проверка успешна, возможны предупреждения;
`1` — диагностика обнаружила проблемы; `2` — ошибка запуска/настроек/аргументов;
`3` — непредвиденная ошибка. `--version` работает без конфигурации.

## Получение писем — этап 2

На компьютере с настроенным классическим Outlook откройте основное окно и выделите
письма в списке. Затем запустите из корня проекта:

```powershell
.\.venv\Scripts\python.exe main.py capture-outlook --config config/default.toml
```

Outlook автоматически не запускается. Для этой команды общая книга не требуется.
Читаются идентификаторы, тема, дата получения и отправитель. Excel-вложения
`.xlsx`, `.xlsm`, `.xls`, `.xlsb` сохраняются в `.runtime/work/<run_id>/mail-NNNN/`.
Формат файла пока определяется по расширению; его содержимое проверяется на этапе 3.
Небезопасные исходные имена никогда не используются для построения пути сохранения.

`manifest.json` содержит данные писем, позиции в исходном выделении, оригинальные
имена вложений, относительные пути, SHA-256 и ошибки. Время получения записывается
в UTC. Адрес Exchange может быть служебным адресом, который вернул Outlook;
он не подменяет поставщика из таблицы. Тело письма не читается и не сохраняется.
Отчёт содержит служебные данные и хранится локально вне Git.

Корректные вложения сохраняются даже при ошибках других вложений. Пустые и
неполностью сохранённые файлы не включаются в результат; временные файлы удаляются,
а невозможность их удаления явно сообщается как ошибка. Повторный запуск получает
свой каталог; предыдущие результаты не перезаписываются.
Это получение исходных данных, а не импорт товаров в общую книгу.

Для `capture-outlook`: `0` — есть сохранённые вложения и нет ошибок;
`1` — частичные ошибки либо нет Excel-вложений; `2` — запуск/Outlook недоступен;
`3` — непредвиденная ошибка. Предупреждения также выводятся при успешном запуске.
Если невозможно сохранить JSON-отчёт, полученные файлы остаются в каталоге запуска.
Предоставленные MSG пока не открываются этой командой: она читает выделение Outlook.

## Настройки сезонов

Каждый `[[routes]]` содержит уникальный `id`, список названий `aliases`,
имя расчётной вкладки `calculation_sheet` и листа `photo_sheet`.
Можно добавлять новые маршруты без изменения моделей и исходного кода.
Неоднозначные названия, повторные назначения вкладок и опечатки в ключах отклоняются.

Настроены Весна/Весна-лето, Пасха, Лето/Лето-осень, Осень, Зима.
В исходной общей книге зимней пары нет: на этом этапе конфигурация только описывает
её, а вкладки не создаются. Создание пары по выбранному шаблону и редактор сезонов
относятся к этапам 6–7.

Перечень `.xlsx`, `.xlsm`, `.xls`, `.xlsb` в настройках — целевые форматы.
Поддержка их чтения и интеграционные испытания выполняются на этапе 3.
Файлы с паролем не обещаются к поддержке без отдельного решения.

## Проверки

```powershell
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe -m ruff check .
.\.venv\Scripts\python.exe -m mypy
```

Тесты проверяют конфигурацию и модели, ошибки запуска, журналы, COM-жизненный цикл,
смешанное выделение, дубли, безопасные имена, частичные ошибки вложений и отчёт.
COM-объекты в автоматических тестах заменены управляемыми тестовыми объектами.
Настоящая интеграция с Outlook ещё не испытана; порядок приёмки описан отдельно.

## Документы

- `docs/specification.md` — согласованные бизнес-правила и открытые вопросы.
- `docs/architecture/decisions.md` — границы слоёв и решения по Office/сохранению.
- `docs/roadmap.md` — этапы и критерии приёмки.
- `docs/stage-1-code.md` — полные тексты кода и настроек этапа 1 для ревью.
- `docs/stage-1-validation.md` — результаты и границы выполненных проверок.
- `docs/stage-2-code.md` — полные актуальные файлы этапа 2 для ревью.
- `docs/stage-2-validation.md` — проверки и приёмка на другом компьютере.

Рабочие письма и книги, логи, резервные копии, `.venv` и личные настройки исключены
из Git. Коммиты и публикация в GitHub выполняются отдельно по запросу.
````

## docs/architecture/decisions.md

````markdown
# Архитектурные решения

## Слои и зависимости

Domain зависит только от стандартной библиотеки: dataclass-модели, инварианты,
происхождение записи, отчёт. Application будет содержать сценарии и порты Protocol.
Infrastructure реализует работу с конфигурацией, Office, файлами и журналированием.
Presentation преобразует пользовательский ввод в вызовы приложения.
`app.py` — composition root; его диагностический сценарий относится к этапу 1.

Каждый адаптер отвечает за одну внешнюю систему. Outlook/Excel не импортируются
в Domain и не передают COM-объекты в модели. Объекты Office живут в потоке с
инициализированным COM; GUI будет получать события через очередь и tkinter.after.

## Формат настроек

TOML читается штатным tomllib Python 3.12. Загрузка строгая: неизвестные ключи,
неверные типы, конфликты псевдонимов и вкладок приводят к ConfigurationError.
Настройки и маршруты неизменяемы. Сезон — расширяемая строка, не закрытый enum.
Загрузка конфигурации ещё не означает реализацию выбора сезона (этап 4).

## Модели и данные

Артикул — строка, ведущие нули и регистр сохраняются. Текстовые поля не преобразуются
в формулы Excel. ProductRecord содержит копию словаря значений с запретом изменений,
ссылку на письмо/вложение и координаты исходной строки.
Фото имеет хеш нормализованных пикселей; фактическое извлечение относится к этапу 5.
ReceivedTime должен преобразовываться адаптером в datetime с часовым поясом.
Денежные значения могут храниться в Decimal; NaN и бесконечность недопустимы.

## Office и запись

openpyxl предназначен для чтения подходящих OOXML-вложений. Общую книгу с объектами
и нативными изображениями в ячейке будет сохранять Excel через xlwings/pywin32.
Конкретный API изображений и сохранение после повторного открытия проверяются
на Office 2024 и Microsoft 365 до приёмки этапа 5–6.
xlwings pictures.add с anchor само по себе не доказывает режим изображения в ячейке.

Источник формул поставщика импортируется как вычисленные значения; если кеша нет,
потребуется расчёт в Excel с отключёнными макросами и автоматическим обновлением ссылок.
Политика защищённых файлов будет определена по фактическим входам.

Две строки сезонной пары представляют один товар. Индексы дублей учитывают эту пару,
а не объявляют зеркальные строки конфликтом. Изменение сезона — согласованный перенос
обоих представлений с сохранением прежнего фото, если новое не пришло.

Корректные вложения отделяются от ошибочных до записи. План, резервная копия,
проверка состояния исходной книги, запись в рабочую копию и проверка сохранённого
результата должны предшествовать замене общего файла. Гарантии восстановления
подтверждаются fault-injection испытаниями; Excel сам по себе не транзакционная БД.
В этапе 1 книги только проверяются на чтение и никогда не изменяются.

## Логи и доставка

Loguru: UTF-8, ротация по размеру, хранение по сроку. diagnose/backtrace выключены,
чтобы исключения не выводили значения локальных переменных. Содержимое писем не
логируется автоматически. CLI отключает стандартный stderr-sink Loguru; каждый
экземпляр приложения при завершении освобождает только свой файловый sink.

Настоящие письма, книги и резервные копии не включаются в Git. Фиксированные версии
проверенного окружения хранятся отдельно от диапазонов совместимости.
Сборка EXE, пользовательские каталоги, обновления и подпись относятся к этапу 8.

## Этап 2: источник Outlook

MailSource и CaptureManifestStore — порты Application. CaptureMail создаёт отдельный
каталог UUID, вызывает источник и сохраняет манифест. Данные MailCapture содержат
идентичности, позицию в выделении, пути и контрольные суммы; COM-объектов в них нет.

Адаптер подключается через GetActiveObject только к уже запущенному классическому
Outlook. COM инициализируется STA в вызывающем потоке и освобождается в finally.
Состав выделения считывается до обработки вложений. Объекты другого класса
пропускаются с предупреждением. Ошибки отдельных писем/вложений становятся issues.

Дата получения читается из PR_MESSAGE_DELIVERY_TIME (PT_SYSTIME) и сохраняется
как UTC; локальный ReceivedTime не используется для повторной конвертации времени.
Имена сохранённых файлов генерируются по индексам; исходное имя — только метаданные.
Временный файл публикуется после проверки размера и вычисления SHA-256.
Сохранение вложений не означает, что Excel-содержимое уже проверено или импортировано.

JSON-отчёт связывает сохранённые файлы с письмами. Сбой отчёта оставляет вложения
для восстановления и вызывает явную ошибку. Тема, адрес и идентификаторы писем
сохраняются только в локальном манифесте; логи содержат номера позиций и коды ошибок.
Тела писем не читаются. Новый запуск не перезаписывает предыдущий.

Из-за отсутствия настроенного Outlook здесь интеграционная приёмка перенесена
на другое устройство по согласованию с пользователем. Это не блокирует разработку
следующих этапов, но остаётся обязательной частью приёмки приложения.
````

## docs/stage-2-validation.md

````markdown
# Этап 2: получение писем Outlook

Дата: 09.10.2026. Версия приложения 0.2.0, Windows, Python 3.12.8.

## Автоматические проверки

- pytest: 44 passed, включая регрессию этапа 1.
- Ruff: All checks passed.
- Mypy strict: Success, 25 source files.
- Контролируемые ошибки инициализации COM, отсутствия Outlook и основного окна.
- Смешанное/пустое выделение, ошибка отдельного элемента и отсутствующая дата.
- Дубли EntryID/StoreID; одинаковые EntryID разных хранилищ остаются разными письмами.
- UTC-время получения, сохранение оригинальных имён и позиций писем в отчёте.
- Имена с traversal и зарезервированными Windows-словами не становятся путями.
- Пустое/частичное вложение отклоняется; исправные вложения продолжают сохраняться.
- Существующий файл и отчёт не перезаписываются. Ошибка отчёта не удаляет вложения.
- Повторный запуск изолирован новым run_id.
- Команда capture-outlook не требует и не создаёт общую книгу.

Для COM-сценариев использованы управляемые тестовые объекты и monkeypatch.
У пользователя здесь нет настроенного Outlook/учётной записи. Проверка получения
настоящих писем и вложений НЕ выполнена. Новая реализация не считается принятой
на Office 2024 / Microsoft 365 до проверки на целевых устройствах.

## Приёмка на другом компьютере

1. Установить Python 3.12 с tkinter и создать .venv по README. Установить
   requirements-lock.txt для воспроизведения текущего окружения.
2. Запустить классический Outlook с настроенной учётной записью обычным способом,
   с теми же правами пользователя, с которыми запускается приложение.
3. Выделить несколько писем с Excel-вложениями, включая одинаковые имена файлов.
   Выполнить команду capture-outlook из README.
4. Сравнить число и содержимое полученных вложений с выбранными письмами;
   сверить темы, отправителя и время получения с Outlook, учитывая UTC в отчёте.
5. Проверить сохранённые файлы по SHA-256 и убедиться в отсутствии *.part.
6. Повторить запуск: каталог должен отличаться, прежние файлы должны сохраниться.
7. Проверить отсутствие выделения, письмо без Excel и смешанные типы элементов.
   Ошибки должны быть понятными, остальные письма должны обрабатываться.
8. Проверить закрытый Outlook. Приложение должно вернуть код 2 и инструкцию,
   не запускать Outlook автоматически. Письма не отправляются и не изменяются.
9. Выполнить испытания отдельно на Office 2024 и Microsoft 365. Записать версии
   Windows/Office, разрядность, результаты и при необходимости приложить журнал.

Исходная общая книга не участвует в этапе 2. Предоставленные MSG потребуется
разобрать для подготовки реальных Excel-файлов к испытаниям следующих этапов;
команда capture-outlook не реализует офлайн-чтение MSG.

## Ссылки на контракт Office

- [Selection.Item: выделение может содержать разные типы](https://learn.microsoft.com/en-us/office/vba/api/outlook.selection.item).
- [PropertyAccessor.GetProperty: PT_SYSTIME и преобразование UTC](https://learn.microsoft.com/en-us/office/vba/api/outlook.propertyaccessor.getproperty).
````
