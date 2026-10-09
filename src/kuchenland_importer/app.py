"""Composition root for diagnostics and Outlook source capture."""

from dataclasses import dataclass
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Self

from loguru import logger

from kuchenland_importer.application.capture_mail import CaptureMail, CaptureOutcome
from kuchenland_importer.infrastructure.capture_manifest import JsonCaptureManifestStore
from kuchenland_importer.infrastructure.config_loader import load_settings
from kuchenland_importer.infrastructure.logging import LoggingSession, configure_logging
from kuchenland_importer.infrastructure.paths import RuntimePaths
from kuchenland_importer.infrastructure.settings import Settings


@dataclass(frozen=True, slots=True)
class DiagnosticResult:
    checks: tuple[str, ...]
    warnings: tuple[str, ...]
    errors: tuple[str, ...]


@dataclass(slots=True)
class Application:
    settings: Settings
    paths: RuntimePaths
    logging: LoggingSession

    @classmethod
    def create(
        cls, config_path: Path, *, workbook: Path | None = None, data_dir: Path | None = None
    ) -> Self:
        settings = load_settings(config_path, workbook=workbook, data_dir=data_dir)
        paths = RuntimePaths.prepare(settings.data_dir)
        session = configure_logging(paths.logs, settings)
        logger.info("Приложение запущено; настроено сезонов: {}", len(settings.routes))
        return cls(settings, paths, session)

    def diagnose(self) -> DiagnosticResult:
        checks = ["Настройки корректны.", "Рабочие каталоги доступны для записи."]
        warnings: list[str] = []
        errors: list[str] = []
        for package in ("pywin32", "xlwings", "openpyxl", "pandas", "Pillow", "loguru"):
            try:
                checks.append(f"{package}: {version(package)}")
            except PackageNotFoundError:
                errors.append(f"Не установлена зависимость {package}; установите requirements.txt.")
        workbook = self.settings.workbook
        if workbook is None:
            warnings.append("Общий файл не выбран; задайте application.workbook или --workbook.")
        else:
            try:
                if not workbook.is_file():
                    errors.append(f"Общий файл не найден: {workbook}")
                elif workbook.suffix.lower() not in self.settings.input_extensions:
                    errors.append(f"Неподдерживаемое расширение общего файла: {workbook.suffix}")
                else:
                    with workbook.open("rb") as stream:
                        stream.read(1)
                    checks.append("Общий файл доступен для чтения; запись не проверялась.")
            except OSError as error:
                errors.append(f"Общий файл недоступен: {error}")
        warnings.append("Диагностика не подключается к Office; импорт товаров ещё не реализован.")
        logger.info("Диагностика: ошибок {}, предупреждений {}", len(errors), len(warnings))
        return DiagnosticResult(tuple(checks), tuple(warnings), tuple(errors))

    def capture_outlook(self) -> CaptureOutcome:
        from kuchenland_importer.infrastructure.outlook.source import OutlookMailSource

        use_case = CaptureMail(OutlookMailSource(), JsonCaptureManifestStore())
        return use_case.execute(self.paths.work, self.settings.input_extensions)

    def close(self) -> None:
        logger.info("Приложение завершено.")
        self.logging.close()
