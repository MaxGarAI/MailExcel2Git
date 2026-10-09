import json
from pathlib import Path

import pytest

from kuchenland_importer.infrastructure.file_transaction import WorkbookTransaction, file_hash


def test_failed_writer_preserves_original_and_verified_backup(tmp_path: Path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"original")
    transaction = WorkbookTransaction(target, tmp_path / "backup")
    with pytest.raises(RuntimeError), transaction:
        transaction.work.write_bytes(b"broken")
        raise RuntimeError("writer failed")
    assert target.read_bytes() == b"original"
    assert transaction.backup.read_bytes() == b"original"
    assert not transaction.work.exists() and not transaction.lock.exists()


def test_commit_replaces_file_and_journal_proves_hashes(tmp_path: Path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"original")
    with WorkbookTransaction(target, tmp_path / "backup") as transaction:
        transaction.work.write_bytes(b"new")
        transaction.commit()
    journal = json.loads(transaction.journal.read_text(encoding="utf-8"))
    assert target.read_bytes() == b"new"
    assert journal["state"] == "committed"
    assert file_hash(target) == journal["new_sha256"]
    assert file_hash(transaction.backup) == journal["original_sha256"]


def test_concurrent_master_modification_aborts_commit(tmp_path: Path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"original")
    with WorkbookTransaction(target, tmp_path / "backup") as transaction:
        transaction.work.write_bytes(b"new")
        target.write_bytes(b"manual change")
        with pytest.raises(ValueError, match="изменилась"):
            transaction.commit()
    assert target.read_bytes() == b"manual change"


def test_second_run_does_not_remove_first_runs_lock(tmp_path: Path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"original")
    with WorkbookTransaction(target, tmp_path / "backup") as first:
        with pytest.raises(FileExistsError):
            with WorkbookTransaction(target, tmp_path / "backup"):
                pass
        assert first.lock.exists()


def test_open_excel_owner_file_blocks_import(tmp_path: Path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"original")
    (tmp_path / "~$master.xlsx").write_bytes(b"owner")
    with pytest.raises(ValueError, match="Закройте"):
        with WorkbookTransaction(target, tmp_path / "backup"):
            pass
    assert not target.with_name(target.name + ".import.lock").exists()


def test_failed_atomic_replace_preserves_original(tmp_path: Path, monkeypatch):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"original")
    import kuchenland_importer.infrastructure.file_transaction as module

    real_replace = module.os.replace

    def fail_work_replace(source, destination):
        if destination == target:
            raise PermissionError("locked between validation and replacement")
        return real_replace(source, destination)

    with WorkbookTransaction(target, tmp_path / "backup") as transaction:
        transaction.work.write_bytes(b"new")
        monkeypatch.setattr(module.os, "replace", fail_work_replace)
        with pytest.raises(PermissionError):
            transaction.commit()
    assert target.read_bytes() == b"original"
