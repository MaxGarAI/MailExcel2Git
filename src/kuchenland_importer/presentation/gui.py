"""Graphical entry point with actionable startup diagnostics."""

import argparse
import sys
import tkinter as tk
from collections.abc import Sequence
from pathlib import Path
from tkinter import messagebox
from types import TracebackType

from loguru import logger

from kuchenland_importer.app import Application
from kuchenland_importer.infrastructure.bootstrap import resolve_configuration
from kuchenland_importer.presentation.gui_window import ImportWindow


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Kuchenland — импорт закупок")
    parser.add_argument("--config", type=Path)
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--data-dir", type=Path)
    parser.add_argument("--check-startup", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    root = tk.Tk()
    root.withdraw()
    app: Application | None = None
    logger.remove()
    try:
        if sys.version_info[:2] != (3, 12):
            raise RuntimeError("Требуется Python 3.12.")
        paths = resolve_configuration(args.config)
        config, personal = paths.active, paths.personal
        app = Application.create(config, workbook=args.workbook, data_dir=args.data_dir)

        def callback_error(
            kind: type[BaseException], value: BaseException, trace: TracebackType | None
        ) -> None:
            logger.opt(exception=(kind, value, trace)).error("Ошибка обработчика интерфейса")
            messagebox.showerror(
                "Ошибка интерфейса", f"{value}\nЖурнал: {app.paths.logs}", parent=root
            )

        root.report_callback_exception = callback_error
        window = ImportWindow(root, app, config, personal)
        root.deiconify()
        if args.check_startup:
            from kuchenland_importer.presentation.gui_startup_check import startup_check
            root.update()
            startup_check(root, window, args.check_startup)
        else:
            root.mainloop()
        return 0
    except Exception as error:
        if app:
            logger.exception("Ошибка графического интерфейса")
        if not args.check_startup:
            messagebox.showerror("Kuchenland — запуск", str(error), parent=root)
        return 2
    finally:
        if app:
            app.close()
        try:
            root.destroy()
        except tk.TclError:
            pass
