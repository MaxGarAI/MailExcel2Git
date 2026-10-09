"""Backup and same-directory atomic replacement; a failed writer never touches the original."""

import hashlib
import json
import os
import shutil
from pathlib import Path
from types import TracebackType
from uuid import uuid4

from loguru import logger


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class WorkbookTransaction:
    def __init__(self, target: Path, backup_directory: Path) -> None:
        self.target = target.resolve()
        self.run_id = uuid4().hex
        self.lock = self.target.with_name(self.target.name + ".import.lock")
        self.work = self.target.with_name(f".kl-{self.run_id}{self.target.suffix}")
        self.backup = backup_directory / f"{self.run_id}{self.target.suffix}"
        self.journal = backup_directory / f"{self.run_id}.json"
        self.original_hash = ""
        self.committed = False
        self._locked = False

    def __enter__(self) -> "WorkbookTransaction":
        if self.target.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Общая книга должна быть XLSX или XLSM.")
        try:
            with self.lock.open("x", encoding="utf-8") as stream:
                self._locked = True
                json.dump({"pid": os.getpid(), "run_id": self.run_id}, stream)
            self._check_available()
            self.backup.parent.mkdir(parents=True, exist_ok=True)
            self.original_hash = file_hash(self.target)
            shutil.copy2(self.target, self.backup)
            with self.backup.open("rb+") as stream:
                os.fsync(stream.fileno())
            if file_hash(self.backup) != self.original_hash:
                raise ValueError("Резервная копия не прошла проверку.")
            shutil.copy2(self.backup, self.work)
            self._journal("prepared")
            return self
        except Exception:
            self._cleanup()
            raise

    def _check_available(self) -> None:
        if self.target.with_name("~$" + self.target.name).exists():
            raise ValueError("Закройте общую книгу в Excel перед импортом.")
        if os.name == "nt":
            import ctypes
            from ctypes import wintypes

            api = ctypes.WinDLL("kernel32", use_last_error=True)
            api.CreateFileW.argtypes = [
                wintypes.LPCWSTR,
                wintypes.DWORD,
                wintypes.DWORD,
                ctypes.c_void_p,
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.HANDLE,
            ]
            api.CreateFileW.restype = wintypes.HANDLE
            api.CloseHandle.argtypes = [wintypes.HANDLE]
            handle = api.CreateFileW(str(self.target), 0xC0000000, 0, None, 3, 0, None)
            if handle == ctypes.c_void_p(-1).value:
                raise OSError(ctypes.get_last_error(), "Общая книга занята или недоступна")
            api.CloseHandle(handle)

    def _journal(self, state: str, new_hash: str | None = None) -> None:
        data = {
            "schema_version": 1,
            "run_id": self.run_id,
            "state": state,
            "target": str(self.target),
            "backup": str(self.backup),
            "original_sha256": self.original_hash,
            "new_sha256": new_hash,
        }
        partial = self.journal.with_suffix(".part")
        with partial.open("w", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(partial, self.journal)

    def commit(self) -> None:
        if self.committed:
            raise ValueError("Транзакция уже завершена.")
        self._check_available()
        if file_hash(self.target) != self.original_hash:
            raise ValueError("Общая книга изменилась во время импорта; запись отменена.")
        with self.work.open("rb+") as stream:
            stream.flush()
            os.fsync(stream.fileno())
        new_hash = file_hash(self.work)
        self._journal("ready", new_hash)
        os.replace(self.work, self.target)
        self.committed = True
        # A ready journal plus matching target hash also proves a committed transaction.
        try:
            self._journal("committed", new_hash)
        except OSError:
            pass

    def _cleanup(self) -> None:
        for path in (self.work, self.lock if self._locked else None):
            if path is None:
                continue
            try:
                path.unlink(missing_ok=True)
            except OSError as error:
                logger.warning(
                    "Не удалён файл транзакции {}; сохранено={}: {}", path, self.committed, error
                )
        self._locked = False

    def __exit__(
        self,
        kind: type[BaseException] | None,
        value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self._cleanup()
