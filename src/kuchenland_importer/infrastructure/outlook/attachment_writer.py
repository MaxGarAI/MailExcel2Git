"""Save under generated names, hash actual bytes, publish only complete files."""

from hashlib import file_digest
from pathlib import Path, PureWindowsPath

import pywintypes

from kuchenland_importer.domain.errors import AttachmentError
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.infrastructure.outlook.com_types import ComAttachment


def attachment_extension(name: str) -> str:
    return PureWindowsPath(name).suffix.casefold()


def save_attachment(
    attachment: ComAttachment, mail: MailMetadata, index: int, directory: Path, name: str
) -> SavedAttachment:
    suffix = attachment_extension(name)
    final = directory / f"attachment-{index:04d}{suffix}"
    partial = directory / f"attachment-{index:04d}{suffix}.part"
    if final.exists() or partial.exists():
        raise AttachmentError("Файл с таким индексом уже существует; перезапись запрещена.")
    try:
        attachment.SaveAsFile(str(partial))
        if partial.stat().st_size == 0:
            raise AttachmentError("Outlook сохранил пустое вложение.")
        with partial.open("rb") as stream:
            digest = file_digest(stream, "sha256").hexdigest()
        partial.rename(final)
    except (OSError, pywintypes.com_error, AttachmentError) as error:
        try:
            partial.unlink(missing_ok=True)
        except OSError as cleanup_error:
            raise AttachmentError(
                "Вложение не сохранено; временный файл не удалось удалить."
            ) from cleanup_error
        raise AttachmentError("Не удалось сохранить и проверить вложение.") from error
    return SavedAttachment(mail, index, name, final, digest)
