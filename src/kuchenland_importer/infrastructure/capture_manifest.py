"""UTF-8 source manifest; mail content is never included in diagnostic logs."""

import json
from pathlib import Path

from kuchenland_importer.domain.capture import MailCapture
from kuchenland_importer.domain.errors import CaptureStorageError


class JsonCaptureManifestStore:
    def save(self, capture: MailCapture) -> Path:
        final = capture.directory / "manifest.json"
        partial = capture.directory / "manifest.json.part"
        mail_indices = {(mail.store_id, mail.entry_id): i for i, mail in enumerate(capture.mails)}
        data = {
            "schema_version": 1,
            "stage": "mail_capture",
            "run_id": capture.run_id,
            "selected_count": capture.selected_count,
            "mails": [
                {
                    "selection_index": capture.selection_positions[index],
                    "entry_id": mail.entry_id,
                    "store_id": mail.store_id,
                    "subject": mail.subject,
                    "received_at": mail.received_at.isoformat(),
                    "sender_name": mail.sender_name,
                    "sender_address": mail.sender_address,
                }
                for index, mail in enumerate(capture.mails)
            ],
            "attachments": [
                {
                    "mail_index": mail_indices[(item.mail.store_id, item.mail.entry_id)],
                    "attachment_index": item.attachment_index,
                    "original_name": item.original_name,
                    "path": item.path.relative_to(capture.directory).as_posix(),
                    "sha256": item.sha256,
                }
                for item in capture.attachments
            ],
            "issues": [
                {
                    "code": issue.code,
                    "message": issue.message,
                    "severity": issue.severity.value,
                    "source": issue.source,
                }
                for issue in capture.issues
            ],
        }
        try:
            with partial.open("x", encoding="utf-8") as stream:
                json.dump(data, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
            partial.rename(final)
        except OSError as error:
            raise CaptureStorageError(
                f"Отчёт не сохранён. Полученные файлы оставлены в {capture.directory}."
            ) from error
        return final
