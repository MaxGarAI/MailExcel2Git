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
