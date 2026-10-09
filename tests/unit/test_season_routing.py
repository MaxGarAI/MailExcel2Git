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
