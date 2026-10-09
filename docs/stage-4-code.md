# Этап 4: полные файлы для ревью

Версия 0.4.0. Полные новые и изменённые файлы; рабочие письма, книги и персональные отчёты сюда не включены.

## src/kuchenland_importer/domain/season.py

````python
"""Extensible seasonal destinations and traceable routing decisions."""

import re
import unicodedata
from dataclasses import dataclass

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.report import ImportIssue


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
class SeasonAssignment:
    source_season: str
    source_year: int | None
    effective_year: int | None
    effective_season: str
    route: SeasonRoute
    sales_start_month: int | None
    rule: str
    issues: tuple[ImportIssue, ...] = ()


@dataclass(frozen=True, slots=True)
class RoutedProduct:
    product: ProductRecord
    assignment: SeasonAssignment | None


@dataclass(frozen=True, slots=True)
class RoutedWorkbook:
    products: tuple[RoutedProduct, ...]
    issues: tuple[ImportIssue, ...]
````

## src/kuchenland_importer/application/sales_month.py

````python
"""Read a sales month/year without treating bare numbers as Excel serial dates."""

import re
from datetime import date, datetime

from kuchenland_importer.domain.cell_values import CellValue
from kuchenland_importer.domain.season import normalize_alias

MONTHS = (
    ("январь", "января", "january"),
    ("февраль", "февраля", "february"),
    ("март", "марта", "march"),
    ("апрель", "апреля", "april"),
    ("май", "мая", "may"),
    ("июнь", "июня", "june"),
    ("июль", "июля", "july"),
    ("август", "августа", "august"),
    ("сентябрь", "сентября", "september"),
    ("октябрь", "октября", "october"),
    ("ноябрь", "ноября", "november"),
    ("декабрь", "декабря", "december"),
)


def sales_period(value: CellValue) -> tuple[int, int | None] | None:
    if isinstance(value, date):
        return value.month, value.year
    if not isinstance(value, str) or not value.strip():
        return None
    text = normalize_alias(value)
    for pattern in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%m.%Y", "%Y-%m"):
        try:
            parsed = datetime.strptime(text, pattern)
            return parsed.month, parsed.year
        except ValueError:
            continue
    # A day/month value has no season year; 2000 permits a leap-day value.
    if re.fullmatch(r"\d{1,2}\.\d{1,2}", text):
        try:
            return datetime.strptime(text + ".2000", "%d.%m.%Y").month, None
        except ValueError:
            return None
    match = re.fullmatch(r"([a-zа-я]+)(?:\s+([1-9]\d{3}))?", text)
    if match is not None:
        for month, names in enumerate(MONTHS, 1):
            if match[1] in names:
                return month, int(match[2]) if match[2] else None
    return None


def sales_month(value: CellValue) -> int | None:
    period = sales_period(value)
    return period[0] if period is not None else None
````

## src/kuchenland_importer/application/resolve_season.py

````python
"""Resolve one row using configured aliases and the agreed December exception."""

import re
from collections.abc import Mapping
from types import MappingProxyType

from kuchenland_importer.application.sales_month import sales_period
from kuchenland_importer.domain.errors import ConfigurationError, SeasonRoutingError
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.season import SeasonAssignment, SeasonRoute, normalize_alias


class ResolveSeason:
    def __init__(self, routes: tuple[SeasonRoute, ...]) -> None:
        aliases: dict[str, SeasonRoute] = {}
        by_id: dict[str, SeasonRoute] = {}
        for route in routes:
            if route.id in by_id:
                raise ConfigurationError(f"Сезон {route.id} указан повторно.")
            by_id[route.id] = route
            for alias in route.aliases:
                key = normalize_alias(alias)
                if key in aliases:
                    raise ConfigurationError(f"Название сезона неоднозначно: {alias!r}.")
                aliases[key] = route
        if not routes:
            raise ConfigurationError("Нужно определить хотя бы один сезон.")
        if "winter" in by_id and "spring" not in by_id:
            raise ConfigurationError("Для декабрьского правила нужен маршрут spring.")
        self._aliases: Mapping[str, SeasonRoute] = MappingProxyType(aliases)
        self._routes: Mapping[str, SeasonRoute] = MappingProxyType(by_id)

    def execute(self, product: ProductRecord) -> SeasonAssignment:
        text = normalize_alias(product.season)
        route = self._aliases.get(text)
        year: int | None = None
        if route is None:
            # Only observed service markers and a numeric ordinal are removable.
            text = re.sub(r"^\((?:у|д)\)\s*", "", text)
            text = re.sub(r"^\d+\.\s*", "", text)
            match = re.fullmatch(r"(.+?)\s+([1-9]\d{3})", text)
            if match is not None:
                text, year = match[1], int(match[2])
            route = self._aliases.get(text)
        if route is None:
            raise SeasonRoutingError(
                "SEASON_UNKNOWN",
                f"Не настроен маршрут для сезона {product.season!r}; добавьте его в routes.",
            )
        month, sales_year = sales_period(product.values.get("sales_start")) or (None, None)
        effective_year = year
        rule = "configured_alias"
        issues: tuple[ImportIssue, ...] = ()
        if route.id == "winter":
            if month == 12:
                route = self._routes["spring"]
                rule = "winter_december_to_spring"
                base_year = year if year is not None else sales_year
                if base_year == 9999:
                    raise SeasonRoutingError(
                        "SEASON_YEAR_OUT_OF_RANGE",
                        "Год следующей весны выходит за диапазон 1–9999.",
                    )
                effective_year = base_year + 1 if base_year is not None else None
                if effective_year is None:
                    issues = (
                        ImportIssue(
                            "TARGET_SEASON_YEAR_UNKNOWN",
                            "Назначена Весна; год следующего сезона определить невозможно.",
                            Severity.WARNING,
                            f"{product.source_sheet}:{product.source_row}",
                        ),
                    )
            elif month is None:
                rule = "winter_sales_start_unknown"
                issues = (
                    ImportIssue(
                        "WINTER_SALES_START_UNKNOWN",
                        "Начало продаж не определено; оставлена Зима по согласованному правилу.",
                        Severity.WARNING,
                        f"{product.source_sheet}:{product.source_row}",
                    ),
                )
        return SeasonAssignment(
            source_season=product.season,
            source_year=year,
            effective_year=effective_year,
            effective_season=route.aliases[0]
            + (f" {effective_year}" if effective_year is not None else ""),
            route=route,
            sales_start_month=month,
            rule=rule,
            issues=issues,
        )
````

## src/kuchenland_importer/application/route_workbook.py

````python
"""Keep row diagnostics while making attachment eligibility an all-or-nothing decision."""

from dataclasses import dataclass

from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.domain.errors import SeasonRoutingError
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.season import RoutedProduct, RoutedWorkbook
from kuchenland_importer.domain.workbook import WorkbookPreview


@dataclass(slots=True)
class RouteWorkbook:
    resolver: ResolveSeason

    def execute(self, preview: WorkbookPreview) -> RoutedWorkbook:
        issues = list(preview.issues)
        products = []
        for product in preview.products:
            try:
                assignment = self.resolver.execute(product)
                issues.extend(assignment.issues)
            except SeasonRoutingError as error:
                assignment = None
                issues.append(
                    ImportIssue(
                        error.code,
                        str(error),
                        Severity.ERROR,
                        f"{product.source_sheet}:{product.source_row}",
                    )
                )
            products.append(RoutedProduct(product, assignment))
        return RoutedWorkbook(tuple(products), tuple(issues))
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


class ExcelInputError(ApplicationError):
    """Excel preview inputs or report storage are invalid or unavailable."""


class SeasonRoutingError(ApplicationError):
    """A row has no unambiguous configured seasonal destination."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
````

## src/kuchenland_importer/infrastructure/settings.py

````python
"""Validated immutable configuration; seasons are extensible data, not an enum."""

from dataclasses import dataclass
from pathlib import Path

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.domain.season import SeasonRoute as SeasonRoute
from kuchenland_importer.domain.season import normalize_alias as normalize_alias


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

## src/kuchenland_importer/infrastructure/preview_report.py

````python
"""Serialize typed values explicitly, preserving Excel error types."""

import json
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from kuchenland_importer.application.normalize_workbook import NormalizeWorkbook
from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.application.route_workbook import RouteWorkbook
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.capture_loader import load_attachments
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader


def encode_value(value: object) -> object:
    if isinstance(value, ExcelErrorValue):
        return {"type": "excel_error", "code": value.code}
    if isinstance(value, datetime | date):
        return {"type": "date", "value": value.isoformat()}
    if isinstance(value, Decimal):
        return {"type": "decimal", "value": str(value)}
    raise TypeError(f"Нельзя сериализовать {type(value).__name__}")


@dataclass(frozen=True, slots=True)
class PreviewSummary:
    report: Path
    products: int
    errors: int
    routes_assigned: int = 0
    eligible_products: int = 0


def create_preview(
    manifest: Path, columns: Path, output: Path, *, season_resolver: ResolveSeason | None = None
) -> PreviewSummary:
    try:
        attachments = load_attachments(manifest)
        service = NormalizeWorkbook(load_columns(columns))
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ExcelInputError(f"Не удалось прочитать входные данные: {error}") from error
    result: dict[str, object] = {
        "schema_version": 1,
        "stage": "season_preview" if season_resolver is not None else "excel_preview",
    }
    files = []
    product_count = error_count = 0
    routes_assigned = eligible_products = skipped_attachments = 0
    for attachment in attachments:
        try:
            preview = service.execute(attachment, WorkbookReader().read(attachment.path))
            all_issues = preview.issues
            products = [
                {
                    "article": p.article,
                    "supplier": p.supplier,
                    "season": p.season,
                    "sheet": p.source_sheet,
                    "row": p.source_row,
                    "values": dict(p.values),
                }
                for p in preview.products
            ]
            if season_resolver is not None:
                routed = RouteWorkbook(season_resolver).execute(preview)
                all_issues = routed.issues
                for serialized, item in zip(products, routed.products, strict=True):
                    assignment = item.assignment
                    serialized["routing"] = (
                        {
                            "source_season": assignment.source_season,
                            "source_year": assignment.source_year,
                            "effective_year": assignment.effective_year,
                            "effective_season": assignment.effective_season,
                            "route_id": assignment.route.id,
                            "calculation_sheet": assignment.route.calculation_sheet,
                            "photo_sheet": assignment.route.photo_sheet,
                            "sales_start_month": assignment.sales_start_month,
                            "rule": assignment.rule,
                        }
                        if assignment is not None
                        else None
                    )
                routes_assigned += sum(p.assignment is not None for p in routed.products)
            issues = [
                {
                    "code": i.code,
                    "message": i.message,
                    "severity": i.severity.value,
                    "source": i.source,
                }
                for i in all_issues
            ]
            errors = sum(i.severity.value == "error" for i in all_issues)
            if errors:
                skipped_attachments += 1
            else:
                eligible_products += len(products)
            files.append(
                {
                    "attachment": str(attachment.path),
                    "original_name": attachment.original_name,
                    "sha256": attachment.sha256,
                    "mail_subject": attachment.mail.subject,
                    "received_at": attachment.mail.received_at.isoformat(),
                    "products": products,
                    "issues": issues,
                    "eligible_for_import": errors == 0,
                }
            )
            product_count += len(products)
            error_count += errors
        except Exception as error:
            skipped_attachments += 1
            files.append(
                {
                    "attachment": str(attachment.path),
                    "products": [],
                    "eligible_for_import": False,
                    "issues": [
                        {"code": "WORKBOOK_READ_FAILED", "message": str(error), "severity": "error"}
                    ],
                }
            )
            error_count += 1
    result["files"] = files
    result["product_count"] = product_count
    result["error_count"] = error_count
    if season_resolver is not None:
        result["routes_assigned"] = routes_assigned
        result["eligible_product_count"] = eligible_products
        result["skipped_attachment_count"] = skipped_attachments
    partial = output.with_suffix(".json.part")
    try:
        with partial.open("x", encoding="utf-8") as stream:
            json.dump(
                result, stream, default=encode_value, ensure_ascii=False, indent=2, allow_nan=False
            )
        partial.rename(output)
    except (OSError, ValueError, TypeError) as error:
        raise ExcelInputError(f"Не удалось сохранить отчёт: {error}") from error
    return PreviewSummary(output, product_count, error_count, routes_assigned, eligible_products)
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
    parser = argparse.ArgumentParser(description="Kuchenland: диагностика и предпросмотр импорта")
    parser.add_argument(
        "command",
        nargs="?",
        default="diagnose",
        choices=("diagnose", "capture-outlook", "preview-excel", "preview-seasons"),
    )
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument("--config", type=Path, required=True, help="Путь к настройкам TOML")
    parser.add_argument("--workbook", type=Path, help="Переопределить путь общей книги")
    parser.add_argument("--data-dir", type=Path, help="Переопределить рабочий каталог")
    parser.add_argument("--manifest", type=Path, help="Манифест полученных вложений")
    args = parser.parse_args(argv)
    # The executable owns its logger. Disable the default diagnostic stderr sink.
    logger.remove()
    app: Application | None = None
    try:
        if sys.version_info[:2] != (3, 12):
            print("Требуется Python 3.12.", file=sys.stderr)
            return 2
        app = Application.create(args.config, workbook=args.workbook, data_dir=args.data_dir)
        if args.command in {"preview-excel", "preview-seasons"}:
            from uuid import uuid4

            from kuchenland_importer.application.resolve_season import ResolveSeason
            from kuchenland_importer.infrastructure.preview_report import create_preview

            if args.manifest is None:
                parser.error(f"{args.command} требует --manifest")
            resolver = (
                ResolveSeason(app.settings.routes) if args.command == "preview-seasons" else None
            )
            prefix = "season-preview" if resolver is not None else "excel-preview"
            preview = create_preview(
                args.manifest,
                args.config.resolve().parent / "columns.toml",
                app.paths.reports / f"{prefix}-{uuid4().hex}.json",
                season_resolver=resolver,
            )
            print(f"Прочитано товаров: {preview.products}; ошибок: {preview.errors}")
            if resolver is not None:
                print(f"Назначено маршрутов: {preview.routes_assigned}")
                print(f"Товаров в корректных вложениях: {preview.eligible_products}")
                logger.info(
                    "Предпросмотр сезонов: товаров {}, маршрутов {}, ошибок {}",
                    preview.products, preview.routes_assigned, preview.errors,
                )
            print(f"Отчёт: {preview.report}")
            print("Это предварительный разбор; общая книга не изменялась.")
            return 1 if preview.errors else 0
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

## src/kuchenland_importer/__init__.py

````python
"""Kuchenland procurement importer."""

__version__ = "0.4.0"
````

## tests/unit/test_season_routing.py

````python
from dataclasses import replace
from datetime import UTC, date, datetime
from pathlib import Path

import pytest

from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.application.route_workbook import RouteWorkbook
from kuchenland_importer.application.sales_month import sales_month, sales_period
from kuchenland_importer.domain.cell_values import CellValue, ExcelErrorValue
from kuchenland_importer.domain.errors import ConfigurationError, SeasonRoutingError
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.season import SeasonRoute
from kuchenland_importer.domain.workbook import WorkbookPreview
from kuchenland_importer.infrastructure.config_loader import load_settings

CONFIG = Path(__file__).resolve().parents[2] / "config" / "default.toml"


def product(season: str, sales_start: CellValue = None) -> ProductRecord:
    mail = MailMetadata("entry", "store", "Subject", datetime(2026, 10, 9, tzinfo=UTC), "", "")
    attachment = SavedAttachment(mail, 1, "source.xlsx", CONFIG.parent / "source.xlsx", "0" * 64)
    return ProductRecord(
        "001-A", "Vendor", season, attachment, "Товары", 10,
        {"season": season, "sales_start": sales_start, "custom": "unchanged"},
    )


@pytest.fixture
def resolver() -> ResolveSeason:
    return ResolveSeason(load_settings(CONFIG).routes)


@pytest.mark.parametrize(
    ("name", "destination", "year"),
    [
        ("Весна", "spring", None),
        ("Весна-лето", "spring", None),
        ("Пасха", "easter", None),
        ("ЛЕТО", "summer", None),
        ("ЛЕТО-ОСЕНЬ", "summer", None),
        ("ОСЕНЬ", "autumn", None),
        ("Зима", "winter", None),
        (" (У) 2. ВЕСНА – ЛЕТО 2027 ", "spring", 2027),
        ("(Д) 6. ЗИМА 2026", "winter", 2026),
        ("8. ПАСХА 2027", "easter", 2027),
        ("10. ЛЕТО-ОСЕНЬ 2027", "summer", 2027),
        ("1.Весна 2027", "spring", 2027),
    ],
)
def test_routes_are_independent_of_case_markers_and_year(
    resolver: ResolveSeason, name: str, destination: str, year: int | None
) -> None:
    source = product(name, date(2026, 10, 1))
    result = resolver.execute(source)
    assert result.route.id == destination
    assert result.source_year == year
    assert result.source_season == name
    assert result.route.calculation_sheet != result.route.photo_sheet
    assert source.season == name and source.values["season"] == name


@pytest.mark.parametrize("month", range(1, 13))
def test_winter_changes_only_in_december(resolver: ResolveSeason, month: int) -> None:
    result = resolver.execute(product("6. ЗИМА 2026", date(2026, month, 1)))
    assert result.route.id == ("spring" if month == 12 else "winter")
    assert result.source_year == 2026
    assert result.effective_year == (2027 if month == 12 else 2026)
    assert not result.issues
    if month == 12:
        assert result.rule == "winter_december_to_spring"
        assert result.effective_season == "Весна 2027"


@pytest.mark.parametrize(
    ("value", "month"),
    [
        (datetime(2026, 12, 10, tzinfo=UTC), 12),
        ("10.12.2026", 12),
        ("10/12/2026", 12),
        ("2026-12-10", 12),
        ("12.2026", 12),
        ("2026-12", 12),
        ("10.12", 12),
        (" ДЕКАБРЬ 2026 ", 12),
        ("December", 12),
        ("ноября", 11),
        ("31.11.2026", None),
        ("ноябрь-декабрь", None),
        ("не ранее декабря 2026", None),
        ("2026-13", None),
        (12, None),
        (46366.0, None),
        (True, None),
        (None, None),
        (ExcelErrorValue("#DIV/0!"), None),
    ],
)
def test_sales_month_is_unambiguous(value: CellValue, month: int | None) -> None:
    assert sales_month(value) == month


@pytest.mark.parametrize("value", [None, "", "по согласованию", ExcelErrorValue("#N/A")])
def test_missing_winter_sales_month_warns_and_preserves_values(
    resolver: ResolveSeason, value: CellValue
) -> None:
    source = product("Зима", value)
    result = resolver.execute(source)
    assert result.route.id == "winter"
    assert result.issues[0].code == "WINTER_SALES_START_UNKNOWN"
    assert result.issues[0].severity.value == "warning"
    assert source.values["sales_start"] == value


@pytest.mark.parametrize(
    "name", ["(Д) 7. НГ 2026", "ВСЕСЕЗОННЫЙ", "Весна/Пасха", "(X) Весна", "Весна 2026/2027"]
)
def test_unknown_and_ambiguous_seasons_are_not_guessed(
    resolver: ResolveSeason, name: str
) -> None:
    with pytest.raises(SeasonRoutingError) as failure:
        resolver.execute(product(name))
    assert failure.value.code == "SEASON_UNKNOWN"


def test_new_route_and_sheet_pair_are_data_not_code() -> None:
    custom = SeasonRoute("new_year", ("Новый год", "НГ"), "Новый год", "ФОТО Новый год")
    routes = load_settings(CONFIG).routes + (custom,)
    assignment = ResolveSeason(routes).execute(product("(Д) 7. НГ 2026"))
    assert assignment.route == custom
    assert assignment.source_year == 2026


def test_exact_configured_alias_has_priority_over_service_prefix_parsing() -> None:
    route = SeasonRoute("collection", ("1. Коллекция 2027",), "Коллекция", "ФОТО коллекция")
    assert ResolveSeason((route,)).execute(product("1. Коллекция 2027")).route == route


def test_registry_rejects_alias_conflict_and_missing_spring() -> None:
    routes = load_settings(CONFIG).routes
    conflict = SeasonRoute("custom", ("ВЕСНА",), "Другая", "ФОТО другая")
    with pytest.raises(ConfigurationError, match="неоднозначно"):
        ResolveSeason(routes + (conflict,))
    with pytest.raises(ConfigurationError, match="spring"):
        ResolveSeason((routes[-1],))


def test_batch_keeps_diagnostics_for_mixed_seasons(resolver: ResolveSeason) -> None:
    first = product("Весна")
    second = replace(product("НГ"), article="B", source_row=11)
    third = replace(product("Зима"), article="C", source_row=12)
    result = RouteWorkbook(resolver).execute(WorkbookPreview((first, second, third), ()))
    assert [p.assignment is not None for p in result.products] == [True, False, True]
    assert [i.code for i in result.issues] == ["SEASON_UNKNOWN", "WINTER_SALES_START_UNKNOWN"]
    assert result.products[0].product is first


def test_december_year_can_come_from_calendar_sales_start(resolver: ResolveSeason) -> None:
    result = resolver.execute(product("Зима", date(2026, 12, 10)))
    assert result.source_year is None
    assert result.effective_year == 2027 and result.effective_season == "Весна 2027"
    assert not result.issues
    assert sales_period("декабрь 2026") == (12, 2026)


def test_december_without_any_year_warns(resolver: ResolveSeason) -> None:
    result = resolver.execute(product("Зима", "декабрь"))
    assert result.route.id == "spring" and result.effective_year is None
    assert result.issues[0].code == "TARGET_SEASON_YEAR_UNKNOWN"


def test_next_year_must_stay_in_calendar_range(resolver: ResolveSeason) -> None:
    with pytest.raises(SeasonRoutingError) as failure:
        resolver.execute(product("Зима 9999", "декабрь"))
    assert failure.value.code == "SEASON_YEAR_OUT_OF_RANGE"
````

## tests/integration/test_season_preview.py

````python
import json
from datetime import UTC, datetime
from hashlib import sha256
from pathlib import Path

from openpyxl import Workbook

from kuchenland_importer.presentation.cli import main

CONFIG = Path(__file__).resolve().parents[2] / "config" / "default.toml"


def test_cli_routes_each_row_and_skips_entire_invalid_attachment(tmp_path: Path) -> None:
    rows = [
        [("A", "Зима 2026", "10.12.2026"), ("B", "Осень 2027", "")],
        [("C", "Весна", ""), ("D", "ВСЕСЕЗОННЫЙ", "")],
        [("E", "Зима", "")],
    ]
    attachments = []
    originals = {}
    for index, products in enumerate(rows, 1):
        path = tmp_path / f"source-{index}.xlsx"
        book = Workbook()
        book.active.append(["SKU", "Supplier", "Season", "Sales Start"])
        for article, season, start in products:
            book.active.append([article, "Vendor", season, start])
        book.save(path)
        originals[path] = path.read_bytes()
        attachments.append({
            "mail_index": 0, "attachment_index": index, "original_name": path.name,
            "path": path.name, "sha256": sha256(originals[path]).hexdigest(),
        })
    source = tmp_path / "manifest.json"
    source.write_text(json.dumps({
        "schema_version": 1, "stage": "mail_capture",
        "mails": [{
            "entry_id": "entry", "store_id": "store", "subject": "Subject",
            "sender_name": "", "sender_address": "",
            "received_at": datetime(2026, 10, 9, tzinfo=UTC).isoformat(),
        }], "attachments": attachments,
    }), encoding="utf-8")
    source_before = source.read_bytes()
    runtime = tmp_path / "runtime"
    assert main([
        "preview-seasons", "--config", str(CONFIG), "--manifest", str(source),
        "--data-dir", str(runtime),
    ]) == 1
    report = next((runtime / "reports").glob("season-preview-*.json"))
    data = json.loads(report.read_text(encoding="utf-8"))
    assert data["stage"] == "season_preview"
    assert data["product_count"] == 5 and data["routes_assigned"] == 4
    assert data["error_count"] == 1 and data["skipped_attachment_count"] == 1
    assert data["eligible_product_count"] == 3  # C is valid but its attachment is skipped.
    first, second, third = data["files"]
    assert first["eligible_for_import"] and third["eligible_for_import"]
    assert not second["eligible_for_import"]
    assert first["products"][0]["routing"]["route_id"] == "spring"
    assert first["products"][0]["season"] == "Зима 2026"
    assert first["products"][0]["routing"]["effective_season"] == "Весна 2027"
    assert first["products"][0]["routing"]["source_year"] == 2026
    assert first["products"][1]["routing"]["route_id"] == "autumn"
    assert second["products"][1]["routing"] is None
    assert third["issues"][0]["code"] == "WINTER_SALES_START_UNKNOWN"
    assert all(path.read_bytes() == before for path, before in originals.items())
    assert source.read_bytes() == source_before
````

## pyproject.toml

````toml
[build-system]
requires = ["setuptools>=75,<83"]
build-backend = "setuptools.build_meta"

[project]
name = "kuchenland-importer"
version = "0.4.0"
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
module = ["pythoncom", "pywintypes", "win32com.*", "openpyxl.*"]
ignore_missing_imports = true
````

## README.md

````markdown
# Kuchenland Importer

Windows-приложение для импорта товаров из выделенных писем классического Outlook
в локальную общую книгу Excel. Python 3.12, Office Professional Plus 2024 / Microsoft 365.

## Текущий статус

Реализованы этапы 1–4: основа, получение писем, разбор Excel и маршрутизация сезонов.
На реальных XLSX-вложениях из 13 MSG прочитано 90 товаров. Автоматические тесты проверены;
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

## Предварительный разбор Excel — этап 3

Передайте манифест, созданный командой capture-outlook:

```powershell
.\.venv\Scripts\python.exe main.py preview-excel --config config/default.toml --manifest "ПОЛНЫЙ_ПУТЬ_К_manifest.json"
```

Словарь столбцов находится в `config/columns.toml` рядом с TOML-настройками приложения.
Артикул/Article/SKU/Item No приводятся к одному полю. Поддерживаются многоуровневые
и объединённые заголовки исследованных шаблонов. Порядок именованных столбцов
не влияет на поиск полей. Нераспознанные поля сохраняются с префиксом source:;
безымянные столбцы получают исходный номер, для их автоматического переноса потребуется
явное сопоставление на этапе 6. Неоднозначные обязательные столбцы не угадываются.

Поставщик и сезон берутся из строки расчётной таблицы. Поля фото-листа соединяются
по артикулу и сохраняются с префиксом photo:, изображения пока не извлекаются.
Числовые артикулы с простым форматом 00000 восстанавливают ведущие нули;
текстовые артикулы сохраняют регистр и нули. Распознаваемые цены становятся Decimal,
даты начала продаж — date; пояснительные текстовые значения сохраняются.

Ошибки Excel остаются типизированными `{type: excel_error, code: ...}`, а не обычным
текстом. Они отражаются предупреждениями и не отменяют строку по согласованному правилу.
Если отсутствует сохранённый результат формулы, строка отмечается ошибкой:
openpyxl не вычисляет формулы. Пустой текстовый результат формулы — допустимое значение.

XLSX/XLSM читаются без сохранения книги. XLS/XLSB используют отдельный экземпляр
Excel, отключают VBA через AutomationSecurity и открывают книгу ReadOnly без обновления
ссылок. Нативная совместимость XLS/XLSB, политики Excel 4.0 macros и особенности
защищённых файлов требуют проверки на целевой машине; листы макросов не поддерживаются.
При отсутствии подходящего Excel файл попадёт в отчёт с ошибкой.

SHA-256 вложений проверяется перед чтением. JSON-отчёт записывается в .runtime/reports;
повреждённая книга не блокирует другие вложения. Счётчик товаров — число прочитанных
записей предпросмотра, не число импортированных товаров. eligible_for_import отражает
валидность чтения; финальная проверка сезона, дублей между письмами и записи ещё впереди.
Исходные файлы и общая книга не изменяются. Лимиты: 50 МБ на OOXML-файл,
100000 строк и 512 столбцов на лист; произвольные шаблоны требуют настройки профиля.

Для QA предоставленных Unicode MSG без Outlook есть `tools/extract_msg_samples.py`.
Это ограниченная утилита исследования материалов, а не универсальный MSG-импортёр.
Она сохраняет вложения и совместимый манифест в новый каталог, читая Windows OLE
только для чтения. Исходные MSG проверяются SHA-256. Команда для предоставленной папки:

```powershell
.\.venv\Scripts\python.exe tools/extract_msg_samples.py "D:\вайбкодин\Почта_сводник" ".runtime\msg-qa-new"
```

## Предпросмотр сезонных маршрутов — этап 4

Команда `preview-seasons` читает исходные вложения из манифеста и для каждой строки
назначает расчётную вкладку и соответствующий лист ФОТО. Для уже извлечённых
предоставленных материалов запуск из корня проекта:

```powershell
.\.venv\Scripts\python.exe main.py preview-seasons --config config/default.toml --manifest .runtime/msg-stage3-verified/manifest.json
```

Для писем Outlook передайте путь к `manifest.json` своего запуска capture-outlook.
Настроенные псевдонимы сопоставляются без учёта регистра, с нормализацией пробелов
и вариантов тире. Поддержан исследованный формат `(У) 2. ВЕСНА-ЛЕТО 2027`:
пометки `(У)`/`(Д)`, порядковый номер и один четырёхзначный год отделяются от названия.
Другие пометки, диапазоны годов и составные сезоны без псевдонима не угадываются.
Точное совпадение с настроенным псевдонимом имеет приоритет перед разбором пометок.

Маршрут с id `winter` при месяце начала продаж 12 меняется на маршрут `spring`.
Эти идентификаторы используются декабрьским правилом; если настроен winter,
нужен spring. Имена вкладок и псевдонимы обоих маршрутов можно менять.
Месяц определяется из date/datetime, календарных дат ДД.ММ.ГГГГ, ДД/ММ/ГГГГ,
ГГГГ-ММ-ДД, месяца с годом ММ.ГГГГ/ГГГГ-ММ, ДД.ММ и полного русского/английского
названия месяца, в том числе с годом. Нераспознанный текст, числа без формата даты,
пустые значения и ошибки Excel не интерпретируются произвольно.
При неопределённом начале продаж Зима остаётся на зимней паре с предупреждением.

НГ и ВСЕСЕЗОННЫЙ сейчас не настроены по решению пользователя. Любой неизвестный
сезон делает всё вложение непригодным для будущего импорта, включая остальные
корректные строки этого вложения. Другие вложения продолжают обрабатываться.
Предпросмотр сохраняет все прочитанные строки и объясняет ошибки.

Отчёт `.runtime/reports/season-preview-<run_id>.json` содержит `routing` каждой
строки: исходный сезон/год, итоговый сезон/год, месяц, правило и точную пару вкладок.
Неопределённый маршрут — null. `routes_assigned` считает назначения, в том числе
в ошибочных вложениях; `eligible_product_count` считает только товары из полностью
корректных вложений. Эти числа не означают запись товаров в общую книгу.
Коды завершения: 0 — нет ошибок, 1 — есть ошибочные вложения/строки,
2 — ошибка входных данных или настроек, 3 — непредвиденный сбой.

Исходная строка и её поля сохраняются без изменения; итоговый сезон находится
в назначении маршрута и должен использоваться будущим адаптером записи.
Исходный год сохраняется отдельно. По согласованному правилу декабрьский переход
повышает его на один: Зима 2026 → Весна 2027. Если в названии года нет, используется
год распознанной даты начала продаж. Если нет обоих, Весна назначается без года
с предупреждением TARGET_SEASON_YEAR_UNKNOWN. Для остальных маршрутов год названия
сохраняется; при его отсутствии год из даты автоматически не добавляется.
Наличие/схема вкладок общей книги пока не проверяются, новые листы не создаются;
это этап 6. Изображения и выбор самого нового письма относятся к этапам 5–6.

## Настройки сезонов

Каждый `[[routes]]` содержит уникальный `id`, список названий `aliases`,
имя расчётной вкладки `calculation_sheet` и листа `photo_sheet`.
Можно добавлять новые маршруты без изменения моделей и исходного кода.
Неоднозначные названия, повторные назначения вкладок и опечатки в ключах отклоняются.

Настроены Весна/Весна-лето, Пасха, Лето/Лето-осень, Осень, Зима.
В исходной общей книге зимней пары нет: на этом этапе конфигурация только описывает
её, а вкладки не создаются. Создание пары по выбранному шаблону и редактор сезонов
относятся к этапам 6–7.

Адаптеры чтения `.xlsx`, `.xlsm`, `.xls`, `.xlsb` реализованы на этапе 3.
Реальные вложения проверены для XLSX; испытания XLS/XLSB через Office ещё впереди.
Файлы с паролем не обещаются к поддержке без отдельного решения.

## Проверки

```powershell
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe -m ruff check .
.\.venv\Scripts\python.exe -m mypy
```

Тесты проверяют конфигурацию и модели, ошибки запуска, журналы, COM-жизненный цикл,
смешанное выделение, дубли, безопасные имена, частичные ошибки вложений и отчёт,
поиск заголовков, нормализацию значений и сохранение ошибок Excel.
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
- `docs/stage-3-code.md` — полные новые и изменённые файлы этапа 3.
- `docs/stage-3-validation.md` — результаты чтения и ограничения проверки.
- `docs/stage-4-code.md` — полные новые и изменённые файлы маршрутизации.
- `docs/stage-4-validation.md` — проверки сезонных правил на тестах и вложениях.

Рабочие письма и книги, логи, резервные копии, `.venv` и личные настройки исключены
из Git. После каждого завершённого этапа и успешных проверок код публикуется
в [MailExcel2Git](https://github.com/MaxGarAI/MailExcel2Git) в ветку
`codex/development`. Порядок публикации закреплён в `AGENTS.md`.
````

## docs/specification.md

````markdown
# Спецификация импорта

Версия 0.4. Уточнения пользователя от 09.10.2026 (Europe/Moscow).

## Согласованные правила

1. Целевая среда: Windows, Python 3.12, Office Professional Plus 2024 и Microsoft 365.
   Для COM-сценария требуется классический Outlook; точные сборки испытываются отдельно.
2. Общая книга локальная. Каждый товар представлен на расчётном сезонном листе
   и соответствующем листе ФОТО. Лист `удалено` не участвует в поиске/изменениях.
3. Предварительно артикул уникален во всей книге как идентификатор товара.
   Его две строки в сезонной паре не считаются двумя разными товарами.
4. Источник поставщика — таблица во вложении. Отправитель письма не подменяет поставщика.
   Дата — дата получения выбранного письма Outlook; тема берётся из этого письма.
5. Сезон определяется на уровне строки товара. Одно вложение может иметь разные сезоны.
   Весна/Весна-лето → Весна, Пасха → Пасха, Лето/Лето-осень → Лето,
   Осень → Осень, Зима → Зима. Зима с началом продаж в декабре → Весна.
6. Сезоны и пары вкладок должны расширяться настройками и впоследствии через GUI.
   Создание новой пары потребует определения шаблона; неизвестный сезон нельзя
   безусловно направлять в произвольную вкладку.
7. Входящие данные вставляются как значения, не как формулы поставщика.
   При обновлении заменяются импортируемые поля; формулы в столбцах, которых нет
   во вложении, сохраняются. Служебные ячейки вне строк товаров сохраняются.
   Физический способ замены должен обеспечивать это правило и сохранность ссылок.
8. Изменившаяся фотография заменяет старую; одинаковая не заменяется;
   при отсутствии новой фотографии прежняя сохраняется. Требуется настоящее
   изображение в ячейке Excel, без молчаливой подмены плавающей картинкой.
9. При нескольких письмах с одним артикулом выигрывает самое новое по ReceivedTime.
   Ошибка в новом вложении не означает разрешение незаметно применить старое письмо.
10. Корректные вложения импортируются, ошибочные пропускаются. Частичный результат
    явно отражается в отчёте. Системный сбой записи требует восстановления книги.
11. При смене сезона товар и фото переносятся на новую пару, старые записи удаляются.
12. `Цена под TARGET PRICE` сохраняется и импортируется при наличии.
13. Целевые форматы вложений: `.xlsx`, `.xlsm`, `.xls`, `.xlsb`.
    Для старых/бинарных форматов требуется Excel-адаптер; макросы не запускаются.
14. По уточнению от 09.10.2026 строки с ошибками Excel импортируются с сохранением
    ошибки в ячейке. Ошибка Excel — отдельный тип значения, не текст и не ноль.
    Отсутствие кеша формулы не считается результатом: требуется расчёт в Excel.
15. НГ и ВСЕСЕЗОННЫЙ пока не маршрутизируются. Неизвестный сезон — ошибка,
    всё содержащее его вложение пропускается, остальные вложения обрабатываются.
16. Зима без определимого начала продаж (пустое значение, ошибка Excel,
    неоднозначный текст) остаётся на зимней паре с предупреждением.
17. При декабрьском переходе год повышается: Зима 2026 → Весна 2027.
    Исходный сезон/год сохраняются отдельно для проверки происхождения данных.

## Факты предварительного анализа

В папке D:\вайбкодин\Почта_сводник обнаружено 13 MSG и общая книга.
В исследованной книге 9 листов, 275 файлов изображений. Расчётные и фото-листы
содержат формулы (в частности VLOOKUP на фото-листах). Пользователь уточнил,
что формулы в столбцах вне импортируемого вложения сохраняются.
Существующие ошибки формул не исправляются как побочный эффект импорта.
Схемы фото-листов не одинаковы; некоторые заголовки повторяются (Коллекция).
Нужны профили и контекст заголовка, а не слепой словарь название → колонка.

Предварительный анализ не заменяет полного разбора вложений и проверки Excel.

## Вопросы для соответствующих этапов

- Если даты писем равны и значения расходятся: требуется явный конфликт,
  пока не утверждено правило приоритета.
- Регистр и допустимая нормализация артикулов; значимость пробелов/дефисов;
  дубли внутри вложения и совпадение артикула у разных поставщиков.
- Составные сезоны вне утверждённых псевдонимов.
- Название поля начала продаж и поставщика в реальных вложениях.
- Что делать с ручными значениями в столбцах, отсутствующих во вложении:
  пользователь явно согласовал сохранение формул, но не всех ручных значений.
- Шаблон новой расчётной/фото-пары; какие формулы и оформление переносить.
- Число пользователей и работа с открытой/несохранённой книгой. До согласования
  запись в открытую пользователем книгу не включается.
- Поведение для нескольких фото на один артикул, группированных изображений,
  повреждённых/защищённых файлов и формул без сохранённых результатов.
````

## docs/architecture/decisions.md

````markdown
# Архитектурные решения

## Слои и зависимости

Domain зависит только от стандартной библиотеки: dataclass-модели, инварианты,
происхождение записи, отчёт. Application содержит сценарии и порты Protocol.
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
Выбор пары реализован отдельным сервисом этапа 4; загрузчик только проверяет настройки.

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
на этапе 3 строка получает ошибку. Источник нужно пересчитать в Excel и получить заново.
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

## Этап 3: чтение Excel и нормализация

WorkbookReader возвращает модели SourceSheet/SourceCell без объектов библиотек.
OOXML читается дважды: формулы и сохранённые значения. Это позволяет различать
пустой результат строковой формулы и отсутствующий кеш. LegacyReader изолирует
COM-чтение XLS/XLSB, открывает отдельный экземпляр Excel и закрывает его в finally.

NormalizeWorkbook и ColumnRegistry работают только с моделями Domain. Словарь
псевдонимов загружается отдельным TOML-адаптером. Заголовки связываются с полями,
неизвестные столбцы сохраняются; фото-описания соединяются по артикулу.
Нестандартные шаблоны требуют явного профиля, а не угадывания обязательных полей.

ExcelErrorValue хранит код ошибки как самостоятельный тип. Он допустим в строке
товара по согласованному правилу; будущий адаптер записи должен восстановить ошибку
Excel, а не записать её строковое название. Decimal и даты также сериализуются
с явным типом. Предпросмотр сохраняет источник, строку, тему и дату письма.

Ошибки чтения отдельной книги не останавливают другие книги. Повреждение манифеста
или нарушение контрольных сумм останавливает предпросмотр до чтения: происхождение
данных должно быть достоверным. Предпросмотр не выбирает сезонную пару и не записывает
общую книгу. Метка eligible_for_import означает только отсутствие ошибок чтения;
окончательную готовность определят следующие этапы.

## Этап 4: маршрутизация строк

SeasonRoute перенесён в Domain; Infrastructure сохраняет совместимый импорт этого
типа для прежнего кода. Application не зависит от настроек Infrastructure.
ResolveSeason строит неизменяемый индекс настроенных псевдонимов, разбирает наблюдавшиеся
пометки/номер/год и возвращает SeasonAssignment с причиной выбора пары.
sales_month отвечает только за распознавание однозначного месяца, без обращения к Office.

RouteWorkbook сохраняет исходные ProductRecord и ошибки всех строк. Неизвестное
назначение представлено явно; одна такая строка исключает всё вложение из кандидатов.
Зимняя строка без определимого месяца допускается с предупреждением согласно решению
пользователя. Модели сезонных назначений не меняют исходные данные и не создают вкладки.

preview-seasons использует тот же адаптер чтения, что preview-excel. Отчёт содержит
исходный сезон и назначение отдельно, чтобы будущий writer использовал итоговый сезон
и точные имена пары, сохраняя происхождение данных. Подсчёт назначений отделён от числа
товаров в пригодных вложениях. Ни один из этих счётчиков не означает сохранение книги.
Год назначения при переходе из декабря повышается на один по уточнению пользователя.
Сначала берётся год из сезона, при его отсутствии — из даты начала продаж.
Если ни один источник года недоступен, назначение Весны сохраняется с предупреждением.
````

## docs/roadmap.md

````markdown
# План и критерии приёмки

На 09.10.2026 реализованы этапы 1–4. Их проверки и ограничения зафиксированы
в docs/stage-N-validation.md; интеграция с настоящим Outlook и legacy Excel
остаётся в приёмке на целевом устройстве. Следующий этап — извлечение фотографий.

1. Основа: конфигурация, модели, logger, CLI-диагностика. Проверки настроек,
   моделей и ошибок запуска; ruff, mypy. Не открывает COM и не изменяет книги.
2. Outlook: выбранные письма классического Outlook, ReceivedTime, безопасное
   сохранение вложений, EntryID/StoreID. Проверка пустого/смешанного выделения,
   кириллицы, одинаковых имён вложений; разбор предоставленных MSG на копиях.
3. Вложения Excel: реальные профили поставщиков, синонимы и контекст заголовков,
   значения формул, типы и неизвестные поля. Проверка переставленных столбцов,
   двух «Коллекция», ведущих нулей и поддерживаемых форматов.
4. Сезон: строковые сезоны, год, декабрьское правило, конфликты, настраиваемые
   пары. Тесты всех согласованных маршрутов; неизвестное значение явно сообщается.
5. Фото: связь с артикулом, хеш пикселей, изменённое/отсутствующее фото,
   нативное изображение в ячейке. Проверка в реальном Excel, после save/reopen.
6. Общая книга: частичный импорт корректных вложений, выбор нового письма,
   две строки товара, перенос сезона, сохранение формул вне импортируемых полей,
   backup и восстановление. Проверка повторного импорта, конфликтов, сбоя записи,
   сортировки и фильтрации фотографий; испытания на копиях пользовательской книги.
7. GUI: импорт, настройки сезонных пар и шаблонов, прогресс, результаты,
   запрет параллельных запусков и определённая отмена. Проверка отзывчивости.
8. EXE: воспроизводимая сборка и версия, каталоги пользователя, инструкция,
   проверка Windows без Python и матрица Office 2024 / Microsoft 365.

К следующему этапу переходить после проверки текущего. Завершение этапа 1 не
означает готовность импорта. Готовность коммерческой версии подтверждается
приёмкой реальных сценариев и восстановлением после сбоев на целевых машинах.
````

## docs/stage-4-validation.md

````markdown
# Этап 4: проверки сезонной маршрутизации

09.10.2026 (Europe/Moscow), Python 3.12.8, Windows, версия 0.4.0.

- Полный набор pytest: 121 passed.
- Ruff: All checks passed.
- Mypy strict: Success, 40 source files.
- Проверены все базовые маршруты и 12 месяцев для зимы, регистр, варианты тире,
  пробелы, пометки (У)/(Д), номера и год; добавление нового маршрута без изменения кода.
- Проверены повышение года при декабрьском переходе, определение его из даты при
  отсутствии в названии, предупреждение при неизвестном годе и граница 9999.
- Проверены отсутствие/неоднозначность начала продаж, ошибки Excel, предупреждение
  зимней строки, конфликты псевдонимов и отсутствие spring для декабрьского правила.
- CLI-проверка смешанных сезонов подтверждает, что неизвестный сезон блокирует всё
  вложение, другие вложения продолжают обрабатываться, источники не меняются.

На 13 предоставленных XLSX-вложениях прочитаны 90 товаров. Назначены 87 маршрутов:
Весна — 65, Пасха — 8, Лето — 3, Осень — 2, Зима — 9. Это назначения для всех
распознанных строк, включая строки из вложений, которые будут пропущены целиком.
9 зимних товаров с декабрьским началом продаж получили маршрут Весна.

3 строки с НГ/ВСЕСЕЗОННЫЙ дали ошибки SEASON_UNKNOWN. Они находятся в 2 вложениях;
эти вложения полностью исключены из кандидатов на будущий импорт. В корректных
вложениях остаются 82 товара. Ни один товар ещё не записан в общую книгу.
Все 32 типизированных значения ошибок Excel сохранились в предпросмотре.
SHA-256 до/после совпал для 27 файлов: 13 исходных MSG, 13 извлечённых вложений
и общей книги.

Локальный проверенный отчёт:
`.runtime/reports/season-preview-194946799e784715b6d910b8a23d365b.json`.
Рабочие данные отчёта не публикуются в GitHub.

Границы проверки: назначения пар не подтверждают наличие и схему вкладок в общей
книге. Зимней пары в предоставленной книге нет; её создание относится к этапу 6.
При декабрьском переходе рассчитан следующий год: все 9 реальных зимних товаров
получили Весна 2027, исходное Зима 2026 сохранено в отчёте.
Другие служебные пометки/диапазоны годов/составные названия
требуют явных псевдонимов или профиля. НГ и ВСЕСЕЗОННЫЙ отклоняются по решению пользователя.
Нативные испытания XLS/XLSB, Outlook и изображений в ячейке ещё впереди.
````
