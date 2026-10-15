import sys
from pathlib import Path
from unittest.mock import patch

from kuchenland_importer.infrastructure.bootstrap import resolve_configuration
from kuchenland_importer.infrastructure.config_loader import load_settings


def test_source_config_is_independent_of_working_directory(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    paths = resolve_configuration()
    assert paths.active.is_file()
    assert paths.active.parent.name == "config"


def test_explicit_config_is_not_replaced_by_user_profile(tmp_path):
    explicit = tmp_path / "other" / "settings.toml"
    paths = resolve_configuration(explicit)
    assert paths.active == explicit.resolve()
    assert paths.personal == explicit.parent / "local.toml"
    assert not explicit.parent.exists()


def test_frozen_seeds_profile_and_preserves_user_edits(tmp_path, monkeypatch):
    bundle = Path("config").resolve().parent
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    with (
        patch.object(sys, "frozen", True, create=True),
        patch.object(sys, "_MEIPASS", str(bundle), create=True),
    ):
        paths = resolve_configuration()
        settings = load_settings(paths.active)
        assert settings.data_dir == tmp_path / "KuchenlandImporter" / "data"
        assert (paths.active.parent / "columns.toml").is_file()
        original = paths.active.read_text(encoding="utf-8") + "\n# personal setting\n"
        paths.active.write_text(original, encoding="utf-8")
        assert resolve_configuration() == paths
        assert paths.active.read_text(encoding="utf-8") == original
        assert not list(paths.active.parent.glob(".*.tmp"))
