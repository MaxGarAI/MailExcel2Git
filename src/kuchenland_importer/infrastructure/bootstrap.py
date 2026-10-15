"""Separate immutable bundled defaults from writable per-user configuration."""

import os
import sys
from dataclasses import dataclass, replace
from pathlib import Path
from uuid import uuid4

from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.settings_store import save_settings


@dataclass(frozen=True, slots=True)
class ConfigurationPaths:
    active: Path
    personal: Path


def _seed(source: Path, destination: Path) -> None:
    if destination.exists():
        return
    temporary = destination.with_name(f".{destination.name}-{uuid4().hex}.tmp")
    try:
        with temporary.open("xb") as stream:
            stream.write(source.read_bytes())
            stream.flush()
            os.fsync(stream.fileno())
        try:
            os.link(temporary, destination)
        except FileExistsError:
            pass
    finally:
        temporary.unlink(missing_ok=True)


def resolve_configuration(explicit: Path | None = None) -> ConfigurationPaths:
    if explicit is not None:
        return ConfigurationPaths(explicit.resolve(), explicit.resolve().parent / "local.toml")
    if getattr(sys, "frozen", False):
        resources = Path(str(vars(sys)["_MEIPASS"])) / "config"
        base = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        config = base / "KuchenlandImporter" / "config"
        config.mkdir(parents=True, exist_ok=True)
        for name in ("columns.toml", "writing.toml"):
            _seed(resources / name, config / name)
        personal = config / "local.toml"
        if not personal.exists():
            settings = replace(
                load_settings(resources / "default.toml"),
                data_dir=(config.parent / "data").resolve(),
            )
            staging = config / f".initial-{uuid4().hex}.toml"
            try:
                save_settings(staging, settings)
                _seed(staging, personal)
            finally:
                staging.unlink(missing_ok=True)
        return ConfigurationPaths(personal, personal)
    config = Path(__file__).resolve().parents[3] / "config"
    personal = config / "local.toml"
    return ConfigurationPaths(personal if personal.exists() else config / "default.toml", personal)
