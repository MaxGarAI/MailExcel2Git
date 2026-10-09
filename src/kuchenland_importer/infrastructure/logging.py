"""Owned Loguru sink with rotation and no captured local-variable values."""

from dataclasses import dataclass
from pathlib import Path

from loguru import logger

from kuchenland_importer.domain.errors import StartupError
from kuchenland_importer.infrastructure.settings import Settings


@dataclass(slots=True)
class LoggingSession:
    sink_id: int | None

    def close(self) -> None:
        if self.sink_id is not None:
            logger.remove(self.sink_id)
            self.sink_id = None


def configure_logging(directory: Path, settings: Settings) -> LoggingSession:
    try:
        sink_id = logger.add(
            directory / "application.log",
            level=settings.log_level,
            rotation=f"{settings.log_rotation_mb} MB",
            retention=f"{settings.log_retention_days} days",
            encoding="utf-8",
            enqueue=False,
            backtrace=False,
            diagnose=False,
            catch=False,
            format=(
                "{time:YYYY-MM-DD HH:mm:ss.SSS ZZ} | {level} | {name}:{function}:{line} | {message}"
            ),
        )
    except (OSError, ValueError) as error:
        raise StartupError(f"Не удалось настроить журнал: {error}") from error
    return LoggingSession(sink_id)
