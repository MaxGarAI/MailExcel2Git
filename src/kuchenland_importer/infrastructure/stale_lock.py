"""Explicit crash recovery: release only the journal's lock after its owner has exited."""

import json
import os
import re
from pathlib import Path

from kuchenland_importer.infrastructure.file_transaction import file_hash


def process_alive(pid: int) -> bool:
    if pid <= 0:
        raise ValueError("Неверный PID блокировки.")
    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
    import ctypes
    from ctypes import wintypes

    api = ctypes.WinDLL("kernel32", use_last_error=True)
    api.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    api.OpenProcess.restype = wintypes.HANDLE
    api.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    api.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = api.OpenProcess(0x1000, False, pid)
    if not handle:
        if ctypes.get_last_error() == 87:
            return False
        raise OSError(ctypes.get_last_error(), "Не удалось проверить владельца блокировки")
    try:
        code = wintypes.DWORD()
        if not api.GetExitCodeProcess(handle, ctypes.byref(code)):
            raise OSError(ctypes.get_last_error(), "Не удалось проверить состояние процесса")
        return code.value == 259
    finally:
        api.CloseHandle(handle)


def release_stale_lock(target: Path, journal: Path) -> None:
    data = json.loads(journal.read_text(encoding="utf-8"))
    if data.get("schema_version") != 1 or Path(data["target"]).resolve() != target.resolve():
        raise ValueError("Журнал не принадлежит общей книге.")
    run_id = data["run_id"]
    if not isinstance(run_id, str) or not re.fullmatch(r"[0-9a-f]{32}", run_id):
        raise ValueError("Неверный идентификатор транзакции.")
    if file_hash(target) not in {data["original_sha256"], data.get("new_sha256")}:
        raise ValueError("Книга изменена вручную; требуется разбор резервной копии.")
    lock = target.with_name(target.name + ".import.lock")
    raw = lock.read_bytes()
    owner = json.loads(raw)
    if owner["run_id"] != run_id or type(owner["pid"]) is not int:
        raise ValueError("Блокировка не принадлежит журналу.")
    if process_alive(owner["pid"]):
        raise ValueError("Владелец блокировки ещё работает; восстановление запрещено.")
    if lock.read_bytes() != raw:
        raise ValueError("Блокировка изменилась во время проверки.")
    lock.unlink()
