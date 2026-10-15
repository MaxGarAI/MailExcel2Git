import time
import tkinter as tk
from pathlib import Path
from threading import Event
from unittest.mock import patch

from kuchenland_importer.app import Application
from kuchenland_importer.domain.report import ImportReport
from kuchenland_importer.infrastructure.import_runner import ImportOutcome
from kuchenland_importer.presentation.gui_jobs import GuiResult
from kuchenland_importer.presentation.gui_settings import SettingsDialog
from kuchenland_importer.presentation.gui_window import ImportWindow


def test_real_tk_stays_responsive_and_renders_background_result(tmp_path):
    root = tk.Tk()
    root.withdraw()
    app = Application.create(Path("config/default.toml"), data_dir=tmp_path / "data")
    release, ready = Event(), Event()
    try:
        window = ImportWindow(root, app, Path("config/default.toml"), tmp_path / "local.toml")

        def task(control):
            control.checkpoint("Чтение")
            ready.set()
            assert release.wait(5)
            return GuiResult(
                ImportOutcome(ImportReport("gui-test", products_added=3), None, None, True)
            )

        window.worker.start(task)
        window._busy(True)
        assert ready.wait(5)
        heartbeat = []
        root.after(1, lambda: heartbeat.append(True))
        deadline = time.monotonic() + 2
        while not heartbeat and time.monotonic() < deadline:
            root.update()
            time.sleep(0.005)
        assert heartbeat and window.worker.busy
        assert str(window.entry.cget("state")) == "disabled"
        with patch(
            "kuchenland_importer.presentation.gui_window.messagebox.askyesno", return_value=True
        ):
            window.close()
        assert window.closing
        release.set()
        deadline = time.monotonic() + 5
        while window.worker.busy and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        assert not window.worker.busy
        assert window.last_result.outcome.report.products_added == 3
    finally:
        release.set()
        if window.worker.thread:
            window.worker.thread.join(5)
        app.close()
        try:
            root.destroy()
        except tk.TclError:
            pass


def test_season_dialog_validates_and_persists_pending_edits(tmp_path):
    root = tk.Tk()
    root.withdraw()
    app = Application.create(Path("config/default.toml"), data_dir=tmp_path / "data")
    applied = []
    try:
        dialog = SettingsDialog(root, app.settings, applied.append)
        dialog.withdraw()
        for variable, value in zip(
            dialog.variables, ("custom", "Новый; Новый-лето", "Новый", "ФОТО новый"), strict=False
        ):
            variable.set(value)
        dialog._save()
        assert applied[0].routes[-1].aliases == ("Новый", "Новый-лето")
        dialog = SettingsDialog(root, app.settings, applied.append)
        dialog.withdraw()
        for variable, value in zip(
            dialog.variables, ("bad", "Весна", "Другой", "ФОТО другой"), strict=False
        ):
            variable.set(value)
        with patch("kuchenland_importer.presentation.gui_settings.messagebox.showerror") as error:
            assert not dialog._update()
            error.assert_called_once()
        assert len(applied) == 1
        dialog.destroy()
    finally:
        app.close()
        root.destroy()
