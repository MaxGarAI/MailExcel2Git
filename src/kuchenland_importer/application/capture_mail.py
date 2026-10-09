"""Collect source files in an isolated run without opening the destination book."""

from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from loguru import logger

from kuchenland_importer.application.ports import CaptureManifestStore, MailSource
from kuchenland_importer.domain.capture import MailCapture
from kuchenland_importer.domain.errors import CaptureStorageError


@dataclass(frozen=True, slots=True)
class CaptureOutcome:
    capture: MailCapture
    manifest: Path


@dataclass(slots=True)
class CaptureMail:
    source: MailSource
    manifests: CaptureManifestStore

    def execute(self, work_dir: Path, extensions: tuple[str, ...]) -> CaptureOutcome:
        run_id = uuid4().hex
        directory = work_dir / run_id
        try:
            directory.mkdir(exist_ok=False)
        except OSError as error:
            raise CaptureStorageError("Не удалось создать каталог запуска.") from error
        logger.info("Получение писем: запуск {}", run_id)
        capture = self.source.capture(run_id, directory, extensions)
        for issue in capture.issues:
            logger.warning("Запуск {}: {} ({})", run_id, issue.code, issue.source)
        manifest = self.manifests.save(capture)
        logger.info(
            "Запуск {}: писем {}, вложений {}, ошибок {}",
            run_id,
            len(capture.mails),
            len(capture.attachments),
            capture.error_count,
        )
        return CaptureOutcome(capture, manifest)
