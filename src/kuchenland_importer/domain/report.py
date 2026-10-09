"""Import counters represent saved results, never merely planned changes."""

from dataclasses import dataclass
from enum import StrEnum


class Severity(StrEnum):
    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True, slots=True)
class ImportIssue:
    code: str
    message: str
    severity: Severity
    source: str

    def __post_init__(self) -> None:
        if not self.code.strip() or not self.message.strip() or not self.source.strip():
            raise ValueError("Ошибка должна иметь код, сообщение и источник.")


@dataclass(frozen=True, slots=True)
class ImportReport:
    run_id: str
    mails_processed: int = 0
    attachments_processed: int = 0
    products_added: int = 0
    products_updated: int = 0
    products_unchanged: int = 0
    products_skipped: int = 0
    photos_added: int = 0
    photos_replaced: int = 0
    issues: tuple[ImportIssue, ...] = ()

    def __post_init__(self) -> None:
        if not self.run_id.strip():
            raise ValueError("Отчёт должен иметь идентификатор запуска.")
        counts = (
            self.mails_processed,
            self.attachments_processed,
            self.products_added,
            self.products_updated,
            self.products_unchanged,
            self.products_skipped,
            self.photos_added,
            self.photos_replaced,
        )
        if any(type(value) is not int or value < 0 for value in counts):
            raise ValueError("Счётчики должны быть целыми неотрицательными числами.")

    @property
    def error_count(self) -> int:
        return sum(issue.severity == Severity.ERROR for issue in self.issues)

    @property
    def warning_count(self) -> int:
        return sum(issue.severity == Severity.WARNING for issue in self.issues)
