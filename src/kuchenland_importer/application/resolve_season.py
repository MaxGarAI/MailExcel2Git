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
