"""Verify a capture manifest before processing its source attachments."""

import json
from datetime import datetime
from hashlib import file_digest
from pathlib import Path

from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment


def load_attachments(manifest: Path) -> tuple[SavedAttachment, ...]:
    root = manifest.resolve().parent
    data = json.loads(manifest.read_text(encoding="utf-8"))
    if data.get("schema_version") != 1 or data.get("stage") != "mail_capture":
        raise ValueError("Неподдерживаемый манифест получения писем.")
    mails = [
        MailMetadata(
            m["entry_id"],
            m["store_id"],
            m["subject"],
            datetime.fromisoformat(m["received_at"]),
            m["sender_name"],
            m["sender_address"],
        )
        for m in data["mails"]
    ]
    result = []
    for item in data["attachments"]:
        path = (root / item["path"]).resolve()
        if not path.is_relative_to(root):
            raise ValueError("Путь вложения выходит за пределы каталога запуска.")
        with path.open("rb") as stream:
            digest = file_digest(stream, "sha256").hexdigest()
        if digest != item["sha256"]:
            raise ValueError("Контрольная сумма вложения не совпадает с манифестом.")
        index = item["mail_index"]
        if type(index) is not int or index < 0 or index >= len(mails):
            raise ValueError("Неверный индекс письма в манифесте.")
        result.append(
            SavedAttachment(
                mails[index], item["attachment_index"], item["original_name"], path, digest
            )
        )
    return tuple(result)
