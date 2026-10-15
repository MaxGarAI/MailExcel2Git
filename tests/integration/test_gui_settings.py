from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

import pytest

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.domain.season import SeasonRoute
from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.settings_store import save_settings


def test_settings_roundtrip_cyrillic_quotes_and_new_season(tmp_path):
    settings = load_settings(Path("config/default.toml"), data_dir=tmp_path / "Данные")
    route = SeasonRoute("new_season", ('Сезон "особый"',), "Новая вкладка", "ФОТО новая")
    settings = replace(
        settings,
        routes=settings.routes + (route,),
        workbook=(tmp_path / 'Книга "1".xlsx').resolve(),
    )
    destination = tmp_path / "config" / "local.toml"
    save_settings(destination, settings)
    assert load_settings(destination) == settings


def test_failed_replace_preserves_previous_settings(tmp_path):
    settings = load_settings(Path("config/default.toml"), data_dir=tmp_path)
    destination = tmp_path / "local.toml"
    save_settings(destination, settings)
    previous = destination.read_bytes()
    with (
        patch(
            "kuchenland_importer.infrastructure.settings_store.os.replace",
            side_effect=OSError("locked"),
        ),
        pytest.raises(OSError),
    ):
        save_settings(destination, replace(settings, log_level="DEBUG"))
    assert destination.read_bytes() == previous
    assert list(tmp_path.glob(".*.tmp")) == []


def test_duplicate_season_alias_refused_before_settings_written(tmp_path):
    settings = load_settings(Path("config/default.toml"), data_dir=tmp_path)
    route = SeasonRoute("duplicate", ("Весна",), "Другой", "ФОТО другой")
    with pytest.raises(ConfigurationError):
        replace(settings, routes=settings.routes + (route,))


def test_column_aliases_roundtrip_and_reject_ambiguity_atomically(tmp_path):
    from kuchenland_importer.infrastructure.column_settings import read_fields, save_fields

    fields = read_fields(Path("config/columns.toml"))
    fields["article"].append("Код поставщика")
    destination = tmp_path / "columns.local.toml"
    save_fields(destination, fields)
    assert read_fields(destination)["article"][-1] == "Код поставщика"
    previous = destination.read_bytes()
    fields["supplier"].append("Код поставщика")
    with pytest.raises(ValueError, match="Неоднозначный"):
        save_fields(destination, fields)
    assert destination.read_bytes() == previous
