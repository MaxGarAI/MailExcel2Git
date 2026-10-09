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
