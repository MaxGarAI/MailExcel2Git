"""Non-interactive package acceptance using the real Tk window and bundled Tcl."""

import json
import tkinter as tk
from pathlib import Path

from kuchenland_importer import __version__
from kuchenland_importer.presentation.gui_window import ImportWindow


def startup_check(root: tk.Tk, window: ImportWindow, destination: Path) -> None:
    root.update_idletasks()
    if root.winfo_width() < 900 or not window.actions:
        raise RuntimeError("Main window layout was not initialized.")
    if any(not button.winfo_ismapped() for button in window.actions):
        raise RuntimeError("A command button is not visible.")
    for button in window.actions:
        if button.winfo_x() + button.winfo_width() > button.master.winfo_width():
            raise RuntimeError(f"Command outside toolbar: {button.cget('text')}")
    with destination.open("x", encoding="utf-8") as stream:
        json.dump(
            {
                "version": __version__,
                "tk": str(root.tk.call("info", "patchlevel")),
                "commands": len(window.actions),
                "width": root.winfo_width(),
                "height": root.winfo_height(),
                "status": "ok",
            },
            stream,
        )
