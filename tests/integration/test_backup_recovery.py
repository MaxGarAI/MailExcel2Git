import json
from pathlib import Path

import pytest

from kuchenland_importer.infrastructure.backup_recovery import restore_backup
from kuchenland_importer.infrastructure.file_transaction import WorkbookTransaction
from kuchenland_importer.infrastructure.stale_lock import release_stale_lock


def committed(tmp_path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"original")
    with WorkbookTransaction(target, tmp_path / "backup") as transaction:
        transaction.work.write_bytes(b"imported")
        transaction.commit()
    return target, transaction


def test_restore_preserves_both_versions(tmp_path: Path):
    target, transaction = committed(tmp_path)
    previous = restore_backup(transaction.journal, target, tmp_path / "backup")
    assert target.read_bytes() == b"original"
    assert previous.read_bytes() == b"imported"
    assert restore_backup(transaction.journal, target, tmp_path / "backup") is None


def test_recovery_refuses_manual_changes(tmp_path: Path):
    target, transaction = committed(tmp_path)
    target.write_bytes(b"manual")
    with pytest.raises(ValueError, match="изменялась"):
        restore_backup(transaction.journal, target, tmp_path / "backup")
    assert target.read_bytes() == b"manual"


def test_recovery_refuses_corrupted_backup(tmp_path: Path):
    target, transaction = committed(tmp_path)
    transaction.backup.write_bytes(b"corrupted")
    with pytest.raises(ValueError, match="повреждена"):
        restore_backup(transaction.journal, target, tmp_path / "backup")
    assert target.read_bytes() == b"imported"


def test_recovery_refuses_wrong_target(tmp_path: Path):
    target, transaction = committed(tmp_path)
    other = tmp_path / "other.xlsx"
    other.write_bytes(b"other")
    with pytest.raises(ValueError, match="принадлежит"):
        restore_backup(transaction.journal, other, tmp_path / "backup")


def test_explicit_stale_lock_recovery_checks_owner(tmp_path: Path, monkeypatch):
    target, transaction = committed(tmp_path)
    transaction.lock.write_text(json.dumps({"run_id": transaction.run_id, "pid": 1234}))
    module = "kuchenland_importer.infrastructure.stale_lock.process_alive"
    monkeypatch.setattr(module, lambda _: True)
    with pytest.raises(ValueError, match="ещё работает"):
        release_stale_lock(target, transaction.journal)
    assert transaction.lock.exists()
    monkeypatch.setattr(module, lambda _: False)
    release_stale_lock(target, transaction.journal)
    assert not transaction.lock.exists()
