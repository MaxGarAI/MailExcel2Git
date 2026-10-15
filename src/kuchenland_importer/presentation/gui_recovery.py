"""Recovery runs on the same serialized worker as import."""

from pathlib import Path
from uuid import uuid4

from kuchenland_importer.app import Application
from kuchenland_importer.application.job_control import JobControl
from kuchenland_importer.domain.errors import ApplicationError
from kuchenland_importer.domain.report import ImportReport
from kuchenland_importer.infrastructure.backup_recovery import restore_backup
from kuchenland_importer.infrastructure.import_runner import ImportOutcome
from kuchenland_importer.infrastructure.stale_lock import release_stale_lock
from kuchenland_importer.presentation.gui_jobs import GuiResult


def recovery_job(
    app: Application, journal: Path, lock_only: bool, control: JobControl
) -> GuiResult:
    workbook = app.settings.workbook
    if workbook is None:
        raise ApplicationError("Общая книга не выбрана.")
    control.begin_write()
    try:
        if lock_only:
            release_stale_lock(workbook, journal)
            backup = None
        else:
            backup = restore_backup(journal, workbook, app.paths.backup)
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ApplicationError(f"Восстановление отменено: {error}") from error
    message = (
        "Аварийная блокировка снята."
        if lock_only
        else (
            "Книга восстановлена из резервной копии."
            if backup
            else "Книга уже соответствует резервной копии."
        )
    )
    return GuiResult(
        ImportOutcome(ImportReport(uuid4().hex), None, backup, backup is not None),
        operation_message=message,
    )
