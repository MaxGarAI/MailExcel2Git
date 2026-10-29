"""Office operations run sequentially outside Tk; results contain plain data only."""

from dataclasses import dataclass
from pathlib import Path

from kuchenland_importer.app import Application
from kuchenland_importer.application.job_control import JobControl
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.domain.report import ImportIssue
from kuchenland_importer.infrastructure.column_settings import column_path
from kuchenland_importer.infrastructure.import_runner import ImportOutcome, run_import
from kuchenland_importer.infrastructure.writing_settings import load_writing_settings


@dataclass(frozen=True, slots=True)
class GuiResult:
    outcome: ImportOutcome
    capture_issues: tuple[ImportIssue, ...] = ()
    capture_manifest: Path | None = None
    operation_message: str | None = None


def import_job(
    app: Application, config: Path, control: JobControl, manifest: Path | None = None
) -> GuiResult:
    control.checkpoint("Проверка настроек импорта")
    policy = load_writing_settings(config.parent / "writing.toml")
    workbook = app.settings.workbook
    if (
        workbook is None
        or not workbook.is_file()
        or workbook.suffix.lower() not in {".xlsx", ".xlsm"}
    ):
        raise ExcelInputError("Выберите существующую общую книгу XLSX или XLSM.")
    issues: tuple[ImportIssue, ...] = ()
    if manifest is None:
        control.checkpoint("Получение выделенных писем из классического Outlook")
        capture = app.capture_outlook()
        issues = capture.capture.issues
        manifest = capture.manifest
        if not capture.capture.attachments:
            details = "\n".join(issue.message for issue in issues)
            raise ExcelInputError(
                f"В выделении нет доступных Excel-вложений. {details}\nОтчёт: {manifest}"
            )
    control.checkpoint("Вложения получены")
    outcome = run_import(
        manifest,
        column_path(config),
        app.settings,
        preserve_manual=policy.preserve_manual_values,
        control=control,
        field_review=control.review_fields,
    )
    return GuiResult(outcome, issues, manifest)
