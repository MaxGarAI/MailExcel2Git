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
from kuchenland_importer.presentation.gui_window import ImportWindow


def main(argv: Sequence[str] | None = None) -> int:
    base = Path(__file__).resolve().parents[3]
    personal = base / "config" / "local.toml"
    parser = argparse.ArgumentParser(description="Kuchenland — импорт закупок")
    parser.add_argument("--config", type=Path)
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--data-dir", type=Path)
    args = parser.parse_args(argv)
    root = tk.Tk()
    root.withdraw()
    app: Application | None = None
    logger.remove()
    try:
        if sys.version_info[:2] != (3, 12):
            raise RuntimeError("Требуется Python 3.12.")
        config = (
            args.config or (personal if personal.is_file() else base / "config" / "default.toml")
        ).resolve()
        if args.config:
            personal = config.parent / "local.toml"
        app = Application.create(config, workbook=args.workbook, data_dir=args.data_dir)

        def callback_error(
            kind: type[BaseException], value: BaseException, trace: TracebackType | None
        ) -> None:
            logger.opt(exception=(kind, value, trace)).error("Ошибка обработчика интерфейса")
            messagebox.showerror(
                "Ошибка интерфейса", f"{value}\nЖурнал: {app.paths.logs}", parent=root
            )

        root.report_callback_exception = callback_error
        ImportWindow(root, app, config, personal)
        root.deiconify()
        root.mainloop()
        return 0
    except Exception as error:
        if app:
            logger.exception("Ошибка графического интерфейса")
        messagebox.showerror("Kuchenland — запуск", str(error), parent=root)
        return 2
    finally:
        if app:
            app.close()
        try:
            root.destroy()
        except tk.TclError:
            pass
