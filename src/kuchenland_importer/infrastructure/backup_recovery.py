"""Restore only a verified backup while refusing to overwrite later manual edits."""

import json
import shutil
from pathlib import Path

from kuchenland_importer.infrastructure.file_transaction import WorkbookTransaction, file_hash


def restore_backup(journal_path: Path, target: Path, backup_directory: Path) -> Path | None:
    data = json.loads(journal_path.read_text(encoding="utf-8"))
    if data.get("schema_version") != 1 or Path(data["target"]).resolve() != target.resolve():
        raise ValueError("Журнал не принадлежит указанной общей книге.")
    backup = Path(data["backup"]).resolve()
    if backup.parent != backup_directory.resolve():
        raise ValueError("Резервная копия находится вне настроенного каталога backup.")
    if file_hash(backup) != data["original_sha256"]:
        raise ValueError("Резервная копия повреждена.")
    current = file_hash(target)
    if current == data["original_sha256"]:
        return None  # The failed transaction already left the original intact.
    if current != data.get("new_sha256"):
        raise ValueError("Книга изменялась после импорта; автоматическое восстановление запрещено.")
    with WorkbookTransaction(target, backup_directory) as transaction:
        if transaction.original_hash != current:
            raise ValueError("Книга изменилась при запуске восстановления.")
        shutil.copy2(backup, transaction.work)
        if file_hash(transaction.work) != data["original_sha256"]:
            raise ValueError("Копия для восстановления не прошла проверку.")
        transaction.commit()
    return transaction.backup
