"""Capture the Explorer selection once; isolate errors to items and attachments."""

from pathlib import Path
from typing import cast

import pywintypes

from kuchenland_importer.domain.capture import MailCapture
from kuchenland_importer.domain.errors import AttachmentError, OutlookError
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.infrastructure.outlook.attachment_writer import (
    attachment_extension,
    save_attachment,
)
from kuchenland_importer.infrastructure.outlook.com_types import (
    MailItem,
    OutlookApplication,
    OutlookItem,
    Selection,
)
from kuchenland_importer.infrastructure.outlook.metadata import read_metadata
from kuchenland_importer.infrastructure.outlook.session import outlook_session

OL_MAIL = 43


def _snapshot(selection: Selection) -> tuple[int, list[tuple[int, object]], list[ImportIssue]]:
    count = selection.Count
    if count == 0:
        raise OutlookError("В Outlook не выделены письма.")
    items: list[tuple[int, object]] = []
    issues: list[ImportIssue] = []
    for index in range(1, count + 1):
        try:
            items.append((index, selection.Item(index)))
        except pywintypes.com_error:
            issues.append(
                ImportIssue(
                    "SELECTION_ITEM_FAILED",
                    "Элемент выделения недоступен.",
                    Severity.ERROR,
                    f"selection:{index}",
                )
            )
    return count, items, issues


def capture_selection(
    selection: Selection, run_id: str, directory: Path, extensions: tuple[str, ...]
) -> MailCapture:
    count, items, issues = _snapshot(selection)
    mails: list[MailMetadata] = []
    positions: list[int] = []
    attachments: list[SavedAttachment] = []
    identities: set[tuple[str, str]] = set()
    for ordinal, item in items:
        source = f"selection:{ordinal}"
        try:
            if cast(OutlookItem, item).Class != OL_MAIL:
                issues.append(
                    ImportIssue(
                        "NOT_MAIL",
                        "Выделенный элемент не является письмом.",
                        Severity.WARNING,
                        source,
                    )
                )
                continue
            mail_item = cast(MailItem, item)
            mail = read_metadata(mail_item)
            identity = (mail.store_id, mail.entry_id)
            if identity in identities:
                issues.append(
                    ImportIssue(
                        "DUPLICATE_MAIL", "Повторное письмо пропущено.", Severity.WARNING, source
                    )
                )
                continue
            identities.add(identity)
            mails.append(mail)
            positions.append(ordinal)
            mail_dir = directory / f"mail-{ordinal:04d}"
            mail_dir.mkdir(exist_ok=False)
            collection = mail_item.Attachments
            attachment_count = collection.Count
        except (pywintypes.com_error, OSError, ValueError, AttributeError):
            issues.append(
                ImportIssue(
                    "MAIL_READ_FAILED", "Не удалось прочитать письмо.", Severity.ERROR, source
                )
            )
            continue
        found = 0
        for index in range(1, attachment_count + 1):
            try:
                attachment = collection.Item(index)
                name = attachment.FileName
                if attachment_extension(name) not in extensions:
                    continue
                found += 1
                attachments.append(save_attachment(attachment, mail, index, mail_dir, name))
            except (pywintypes.com_error, OSError, ValueError, AttachmentError):
                issues.append(
                    ImportIssue(
                        "ATTACHMENT_SAVE_FAILED",
                        "Excel-вложение не сохранено.",
                        Severity.ERROR,
                        f"{source}/attachment:{index}",
                    )
                )
        if found == 0:
            issues.append(
                ImportIssue(
                    "NO_EXCEL", "У письма нет доступных Excel-вложений.", Severity.WARNING, source
                )
            )
    return MailCapture(
        run_id, directory, count, tuple(mails), tuple(attachments), tuple(issues), tuple(positions)
    )


class OutlookMailSource:
    def capture(self, run_id: str, directory: Path, extensions: tuple[str, ...]) -> MailCapture:
        with outlook_session() as app:
            try:
                return self._capture_active_explorer(app, run_id, directory, extensions)
            finally:
                del app

    @staticmethod
    def _capture_active_explorer(
        app: OutlookApplication, run_id: str, directory: Path, extensions: tuple[str, ...]
    ) -> MailCapture:
        try:
            explorer = app.ActiveExplorer()
            if explorer is None:
                raise OutlookError("Откройте основное окно Outlook и выделите письма в списке.")
            selection = explorer.Selection
        except pywintypes.com_error as error:
            raise OutlookError("Не удалось прочитать выделение в основном окне Outlook.") from error
        try:
            return capture_selection(selection, run_id, directory, extensions)
        except pywintypes.com_error as error:
            raise OutlookError("Выделение Outlook стало недоступно. Повторите запуск.") from error
