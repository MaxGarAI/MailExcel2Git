"""Diagnostics and source collection; Excel import is not exposed yet."""

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from loguru import logger

from kuchenland_importer import __version__
from kuchenland_importer.app import Application
from kuchenland_importer.domain.errors import ApplicationError


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Kuchenland: диагностика и предпросмотр импорта")
    parser.add_argument(
        "command",
        nargs="?",
        default="diagnose",
        choices=(
            "diagnose", "capture-outlook", "preview-excel", "preview-seasons", "preview-photos"
        ),
    )
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument("--config", type=Path, required=True, help="Путь к настройкам TOML")
    parser.add_argument("--workbook", type=Path, help="Переопределить путь общей книги")
    parser.add_argument("--data-dir", type=Path, help="Переопределить рабочий каталог")
    parser.add_argument("--manifest", type=Path, help="Манифест полученных вложений")
    args = parser.parse_args(argv)
    # The executable owns its logger. Disable the default diagnostic stderr sink.
    logger.remove()
    app: Application | None = None
    try:
        if sys.version_info[:2] != (3, 12):
            print("Требуется Python 3.12.", file=sys.stderr)
            return 2
        app = Application.create(args.config, workbook=args.workbook, data_dir=args.data_dir)
        if args.command in {"preview-excel", "preview-seasons", "preview-photos"}:
            from uuid import uuid4

            from kuchenland_importer.application.resolve_season import ResolveSeason
            from kuchenland_importer.infrastructure.preview_report import create_preview

            if args.manifest is None:
                parser.error(f"{args.command} требует --manifest")
            resolver = (
                ResolveSeason(app.settings.routes) if args.command != "preview-excel" else None
            )
            prefix = "season-preview" if resolver is not None else "excel-preview"
            run_id = uuid4().hex
            photos = (
                app.paths.work / f"photo-preview-{run_id}"
                if args.command == "preview-photos" else None
            )
            if photos is not None:
                prefix = "photo-preview"
            preview = create_preview(
                args.manifest,
                args.config.resolve().parent / "columns.toml",
                app.paths.reports / f"{prefix}-{run_id}.json",
                season_resolver=resolver,
                photo_directory=photos,
            )
            print(f"Прочитано товаров: {preview.products}; ошибок: {preview.errors}")
            if photos is not None:
                print(f"Товаров с извлечённым фото: {preview.photos}")
            if resolver is not None:
                print(f"Назначено маршрутов: {preview.routes_assigned}")
                print(f"Товаров в корректных вложениях: {preview.eligible_products}")
                logger.info(
                    "Предпросмотр сезонов: товаров {}, маршрутов {}, ошибок {}",
                    preview.products, preview.routes_assigned, preview.errors,
                )
            print(f"Отчёт: {preview.report}")
            print("Это предварительный разбор; общая книга не изменялась.")
            return 1 if preview.errors else 0
        if args.command == "capture-outlook":
            outcome = app.capture_outlook()
            capture = outcome.capture
            print(f"Получено писем: {len(capture.mails)} из {capture.selected_count} элементов")
            print(f"Сохранено Excel-вложений: {len(capture.attachments)}")
            for issue in capture.issues:
                print(f"{issue.severity.value}: {issue.source}: {issue.message}")
            print(f"Отчёт: {outcome.manifest}")
            print("Получены исходные данные; общая книга не изменялась.")
            return 1 if capture.error_count or not capture.attachments else 0
        result = app.diagnose()
        print(f"Kuchenland Importer {__version__} — диагностика")
        for label, messages in (
            ("OK", result.checks),
            ("ПРЕДУПРЕЖДЕНИЕ", result.warnings),
            ("ОШИБКА", result.errors),
        ):
            for message in messages:
                print(f"{label}: {message}")
        return 1 if result.errors else 0
    except ApplicationError as error:
        if app is not None:
            logger.warning("Ошибка приложения: {}", error)
        print(f"Ошибка запуска: {error}", file=sys.stderr)
        return 2
    except Exception:
        if app is not None:
            logger.exception("Непредвиденная ошибка приложения")
            message = f"Непредвиденная ошибка. Журнал: {app.paths.logs / 'application.log'}"
        else:
            message = "Непредвиденная ошибка до открытия журнала. Проверьте настройки и доступы."
        print(message, file=sys.stderr)
        return 3
    finally:
        if app is not None:
            app.close()
