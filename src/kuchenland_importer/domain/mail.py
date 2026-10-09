"""Mail provenance preserved independently of the Outlook COM object."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True, slots=True)
class MailMetadata:
    entry_id: str
    store_id: str
    subject: str
    received_at: datetime
    sender_name: str
    sender_address: str

    def __post_init__(self) -> None:
        if not self.entry_id.strip() or not self.store_id.strip():
            raise ValueError("Письмо должно иметь EntryID и StoreID.")
        if self.received_at.tzinfo is None or self.received_at.utcoffset() is None:
            raise ValueError("Дата получения письма должна содержать часовой пояс.")


@dataclass(frozen=True, slots=True)
class SavedAttachment:
    mail: MailMetadata
    attachment_index: int
    original_name: str
    path: Path
    sha256: str

    def __post_init__(self) -> None:
        if self.attachment_index < 1:
            raise ValueError("Индекс вложения Outlook начинается с 1.")
        if not self.original_name.strip() or not self.path.is_absolute():
            raise ValueError("Вложение должно иметь имя и абсолютный путь.")
        if len(self.sha256) != 64 or any(c not in "0123456789abcdef" for c in self.sha256):
            raise ValueError("Ожидается SHA-256 вложения в нижнем регистре.")
