from pathlib import Path
from unittest.mock import patch

import pytest

from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.import_runner import run_import


def test_master_read_failure_names_the_master_instead_of_an_attachment(tmp_path):
    target = tmp_path / "Общая книга.xlsx"
    target.write_bytes(b"not an Excel archive")
    settings = load_settings(Path("config/default.toml"), workbook=target, data_dir=tmp_path)
    with patch("kuchenland_importer.infrastructure.import_runner.prepare_batch", return_value=()):
        with pytest.raises(ExcelInputError) as failure:
            run_import(tmp_path / "capture.json", Path("config/columns.toml"), settings)
    message = str(failure.value)
    assert "чтение общей книги «Общая книга.xlsx»" in message
    assert "Вложение" not in message
    assert target.read_bytes() == b"not an Excel archive"
    assert len(list((tmp_path / "backup").glob("*.xlsx"))) == 1
