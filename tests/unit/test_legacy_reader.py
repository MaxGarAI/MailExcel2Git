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


@pytest.mark.parametrize("code,is_error", [(2007, True), (2045, True), (2045, False)])
def test_legacy_error_is_never_imported_as_hresult(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, code: int, is_error: bool
) -> None:
    value = -2146828288 + code
    cell = SimpleNamespace(Value=value, NumberFormat="General", MergeCells=False)
    used = SimpleNamespace(
        Row=1, Column=1, Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=1)
    )
    sheet = SimpleNamespace(Type=-4167, Name="Data", UsedRange=used, Cells=Mock(return_value=cell))
    book = SimpleNamespace(Sheets=[sheet], Worksheets=[sheet], Close=Mock())
    app = SimpleNamespace(
        Workbooks=SimpleNamespace(Open=Mock(return_value=book)),
        Quit=Mock(),
        WorksheetFunction=SimpleNamespace(IsError=Mock(return_value=is_error)),
    )
    monkeypatch.setattr(legacy_reader.pythoncom, "CoInitializeEx", Mock())
    monkeypatch.setattr(legacy_reader.pythoncom, "CoUninitialize", Mock())
    monkeypatch.setattr(legacy_reader.win32com.client, "DispatchEx", Mock(return_value=app))
    if is_error and code == 2045:
        with pytest.raises(ValueError, match="Неподдерживаемая ошибка"):
            legacy_reader.read_legacy(tmp_path / "source.xls")
    else:
        result = legacy_reader.read_legacy(tmp_path / "source.xls")
        expected = legacy_reader.ExcelErrorValue("#DIV/0!") if is_error else value
        assert result[0].rows[0][0].value == expected
    book.Close.assert_called_once_with(SaveChanges=False)
    app.Quit.assert_called_once()
