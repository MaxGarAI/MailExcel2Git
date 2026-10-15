from pathlib import Path
from unittest.mock import patch

import pytest
from openpyxl import Workbook

from kuchenland_importer.application.job_control import JobCancelled, JobControl
from kuchenland_importer.domain.master import MasterSnapshot
from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.file_transaction import file_hash
from kuchenland_importer.infrastructure.import_runner import run_import


def test_cancel_during_master_read_discards_work_copy_and_keeps_original(tmp_path):
    target = tmp_path / "master.xlsx"
    book = Workbook()
    book.active["A1"] = "Original"
    book.save(target)
    original = file_hash(target)
    control = JobControl(lambda event: None)
    settings = load_settings(Path("config/default.toml"), workbook=target, data_dir=tmp_path)

    def read(*args):
        control.request_cancel()
        return MasterSnapshot({}, {}, ())

    with (
        patch("kuchenland_importer.infrastructure.import_runner.prepare_batch", return_value=()),
        patch("kuchenland_importer.infrastructure.import_runner.read_snapshot", side_effect=read),
        patch("kuchenland_importer.infrastructure.import_runner.write_master") as writer,
        pytest.raises(JobCancelled),
    ):
        run_import(
            tmp_path / "manifest.json", Path("config/columns.toml"), settings, control=control
        )
    writer.assert_not_called()
    assert file_hash(target) == original
    assert not list(tmp_path.glob("*.import.lock"))
    assert not list(tmp_path.glob(".kl-*"))
