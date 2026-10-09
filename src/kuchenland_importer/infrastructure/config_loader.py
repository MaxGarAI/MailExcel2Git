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
