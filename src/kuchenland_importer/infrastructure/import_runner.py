"""Orchestrate preparation, planning, Excel write, verification and atomic commit."""

from dataclasses import asdict, dataclass
from pathlib import Path

from loguru import logger

from kuchenland_importer.application.job_control import JobCancelled, JobControl, Progress
from kuchenland_importer.application.plan_import import make_plan
from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.domain.photos import PhotoAction
from kuchenland_importer.domain.report import ImportIssue, ImportReport, Severity
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.com_values import ERROR_FORMULAS
from kuchenland_importer.infrastructure.excel.master_snapshot import read_snapshot
from kuchenland_importer.infrastructure.excel.master_writer import write_master
from kuchenland_importer.infrastructure.excel.verify_master import verify_master
from kuchenland_importer.infrastructure.file_transaction import WorkbookTransaction
from kuchenland_importer.infrastructure.import_report import save_report
from kuchenland_importer.infrastructure.prepare_batch import prepare_batch
from kuchenland_importer.infrastructure.settings import Settings


@dataclass(frozen=True, slots=True)
class ImportOutcome:
    report: ImportReport
    report_path: Path | None
    backup: Path | None
    committed: bool


def run_import(
    manifest: Path,
    columns: Path,
    settings: Settings,
    *,
    preserve_manual: bool = True,
    control: JobControl | None = None,
) -> ImportOutcome:
    if settings.workbook is None:
        raise ExcelInputError("Не указан путь общей книги.")
    transaction = WorkbookTransaction(settings.workbook, settings.data_dir / "backup")
    work = settings.data_dir / "work" / f"import-{transaction.run_id}"
    reports = settings.data_dir / "reports"
    work.mkdir(parents=True, exist_ok=False)
    reports.mkdir(parents=True, exist_ok=True)
    registry = load_columns(columns)
    if control:
        control.checkpoint("Чтение вложений, сезонов и фотографий")
    try:
        batch = prepare_batch(manifest, registry, ResolveSeason(settings.routes), work / "incoming")
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ExcelInputError(f"Не удалось проверить манифест: {error}") from error
    logger.info("Подготовлено вложений: {}; запуск {}", len(batch), transaction.run_id)
    if control:
        control.checkpoint("Проверка общей книги и существующих товаров")
    phase = "создание резервной и рабочей копий общей книги"
    try:
        with transaction:
            phase = f"чтение общей книги «{settings.workbook.name}»"
            master = read_snapshot(
                transaction.work,
                registry,
                settings.routes,
                settings.excluded_sheets,
                work / "previous",
            )
            phase = "проверка маршрутов и новых версий товаров"
            plan = make_plan(batch, master, frozenset(ERROR_FORMULAS))
            if control:
                control.checkpoint("Проверка маршрутов и выбор новых писем")
            if plan.products:
                if control:
                    control.begin_write()
                phase = f"запись общей книги «{settings.workbook.name}»"
                written = write_master(
                    transaction.work,
                    master,
                    plan,
                    preserve_manual=preserve_manual,
                    progress=(
                        lambda done, total: control.publish(
                            Progress(
                                f"Записано товаров: {done} из {total}. Отмена недоступна.", False
                            )
                        )
                    )
                    if control
                    else None,
                )
                if control:
                    control.publish(Progress("Проверка сохранённых значений и фотографий", False))
                phase = f"проверка сохранённой общей книги «{settings.workbook.name}»"
                verify_master(transaction.work, master, plan, written, work / "verified")
                phase = f"замена общей книги «{settings.workbook.name}»"
                transaction.commit()
            else:
                written = ()
            report = ImportReport(
                run_id=transaction.run_id,
                mails_processed=len(
                    {(b.attachment.mail.entry_id, b.attachment.mail.store_id) for b in batch}
                ),
                attachments_processed=sum(b.eligible for b in plan.attachments),
                products_added=sum(p.previous is None for p in plan.products) if written else 0,
                products_updated=sum(p.previous is not None for p in plan.products)
                if written
                else 0,
                products_unchanged=plan.unchanged,
                products_skipped=plan.skipped,
                photos_added=sum(p.photo_action == PhotoAction.ADD for p in written),
                photos_replaced=sum(p.photo_action == PhotoAction.REPLACE for p in written),
                issues=plan.issues,
            )
    except JobCancelled:
        logger.info("Импорт отменён пользователем до записи")
        raise
    except Exception as error:
        logger.exception("Импорт не завершён; резервная копия: {}", transaction.backup)
        failed = ImportReport(
            transaction.run_id,
            issues=(
                ImportIssue(
                    "IMPORT_ABORTED",
                    f"{phase}: {error}",
                    Severity.ERROR,
                    str(settings.workbook),
                ),
            ),
        )
        save_report(
            reports / f"import-{transaction.run_id}.json",
            {
                "schema_version": 1,
                "committed": transaction.committed,
                "backup": str(transaction.backup),
                "journal": str(transaction.journal),
                "report": asdict(failed),
            },
            committed=transaction.committed,
        )
        raise ExcelInputError(
            f"Импорт отменён — {phase}: {error}. Backup: {transaction.backup}"
        ) from error
    report_path = reports / f"import-{transaction.run_id}.json"
    data = {
        "schema_version": 1,
        "committed": transaction.committed,
        "backup": str(transaction.backup),
        "journal": str(transaction.journal),
        "report": asdict(report),
        "products": [asdict(p) for p in written],
        "attachments": [
            {
                "name": p.attachment.original_name,
                "sha256": p.attachment.sha256,
                "eligible": p.eligible,
                "articles": sorted(p.articles),
                "ignored_products": [asdict(row) for row in p.ignored],
                "issues": [asdict(issue) for issue in p.issues],
            }
            for p in plan.attachments
        ],
    }
    report_saved = save_report(report_path, data, committed=transaction.committed)
    logger.info(
        "Импорт: добавлено {}, обновлено {}, ошибок {}; сохранено={}",
        report.products_added,
        report.products_updated,
        report.error_count,
        transaction.committed,
    )
    return ImportOutcome(
        report, report_path if report_saved else None, transaction.backup, transaction.committed
    )
