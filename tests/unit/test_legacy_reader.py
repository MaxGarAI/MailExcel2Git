from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from kuchenland_importer.infrastructure.excel import legacy_reader


def test_owned_excel_is_closed_after_open_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    app = SimpleNamespace(
        Workbooks=SimpleNamespace(Open=Mock(side_effect=RuntimeError("failed"))), Quit=Mock()
    )
    initialize, uninitialize = Mock(), Mock()
    monkeypatch.setattr(legacy_reader.pythoncom, "CoInitializeEx", initialize)
    monkeypatch.setattr(legacy_reader.pythoncom, "CoUninitialize", uninitialize)
    monkeypatch.setattr(legacy_reader.win32com.client, "DispatchEx", Mock(return_value=app))
    with pytest.raises(RuntimeError, match="failed"):
        legacy_reader.read_legacy(tmp_path / "source.xls")
    assert app.AutomationSecurity == 3
    assert app.EnableEvents is False
    assert app.Workbooks.Open.call_args.kwargs["UpdateLinks"] == 0
    assert app.Workbooks.Open.call_args.kwargs["ReadOnly"] is True
    app.Quit.assert_called_once()
    uninitialize.assert_called_once()


def test_legacy_book_never_saved(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    book = SimpleNamespace(Sheets=[], Worksheets=[], Close=Mock())
    app = SimpleNamespace(Workbooks=SimpleNamespace(Open=Mock(return_value=book)), Quit=Mock())
    monkeypatch.setattr(legacy_reader.pythoncom, "CoInitializeEx", Mock())
    monkeypatch.setattr(legacy_reader.pythoncom, "CoUninitialize", Mock())
    monkeypatch.setattr(legacy_reader.win32com.client, "DispatchEx", Mock(return_value=app))
    assert legacy_reader.read_legacy(tmp_path / "source.xlsb") == ()
    book.Close.assert_called_once_with(SaveChanges=False)
    app.Quit.assert_called_once()
