from pathlib import Path

import pytest

from kuchenland_importer.domain.errors import ConfigurationError
from kuchenland_importer.infrastructure.writing_settings import load_writing_settings


def test_agreed_write_policy_is_loaded():
    path = Path(__file__).resolve().parents[2] / "config/writing.toml"
    result = load_writing_settings(path)
    assert result.preserve_manual_values
    assert result.unknown_fields == "skip"
    assert result.missing_sheets == "skip_attachment"


@pytest.mark.parametrize(
    "change",
    [
        ("preserve_manual_values = true", "preserve_manual_values = 1"),
        ('unknown_fields = "skip"', 'unknown_fields = "create"'),
        ('missing_sheets = "skip_attachment"', 'missing_sheets = "create"'),
        ("schema_version = 1", "schema_version = true"),
    ],
)
def test_invalid_policy_does_not_silently_change_import_behavior(tmp_path, change):
    source = Path(__file__).resolve().parents[2] / "config/writing.toml"
    config = tmp_path / "writing.toml"
    config.write_text(source.read_text(encoding="utf-8").replace(*change), encoding="utf-8")
    with pytest.raises(ConfigurationError):
        load_writing_settings(config)
