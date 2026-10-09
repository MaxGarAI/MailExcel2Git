"""Result of collecting source mail; no Excel rows have been imported yet."""

from dataclasses import dataclass
from pathlib import Path

from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.report import ImportIssue, Severity


@dataclass(frozen=True, slots=True)
class MailCapture:
    run_id: str
    directory: Path
    selected_count: int
    mails: tuple[MailMetadata, ...]
    attachments: tuple[SavedAttachment, ...]
    issues: tuple[ImportIssue, ...]
    selection_positions: tuple[int, ...]

    def __post_init__(self) -> None:
        if not self.run_id.strip() or not self.directory.is_absolute():
            raise ValueError("Запуск должен иметь идентификатор и абсолютный каталог.")
        if type(self.selected_count) is not int or self.selected_count < len(self.mails):
            raise ValueError("Число элементов выделения не может быть меньше числа писем.")
        identities = {(mail.store_id, mail.entry_id) for mail in self.mails}
        if len(identities) != len(self.mails):
            raise ValueError("В результате получения писем обнаружены дубли.")
        if len(self.selection_positions) != len(self.mails) or any(
            index < 1 or index > self.selected_count for index in self.selection_positions
        ):
            raise ValueError("Позиции писем должны соответствовать исходному выделению.")
        for attachment in self.attachments:
            if (attachment.mail.store_id, attachment.mail.entry_id) not in identities:
                raise ValueError("Вложение не связано с полученным письмом.")
            if not attachment.path.is_relative_to(self.directory):
                raise ValueError("Вложение находится вне каталога запуска.")

    @property
    def error_count(self) -> int:
        return sum(issue.severity == Severity.ERROR for issue in self.issues)
