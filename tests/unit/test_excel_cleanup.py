from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from kuchenland_importer.infrastructure.excel import owned_session


@pytest.mark.parametrize("primary_failure", [True, False])
def test_excel_cleanup_preserves_primary_error_and_never_hides_cleanup_only_failure(
    monkeypatch, primary_failure
):
    seed = SimpleNamespace(Close=Mock())
    book = SimpleNamespace(
        ReadOnly=False,
        ProtectStructure=False,
        Sheets=[],
        Close=Mock(side_effect=RuntimeError("close failed")),
    )
    app = SimpleNamespace(
        Workbooks=SimpleNamespace(Add=Mock(return_value=seed), Open=Mock(return_value=book)),
        Quit=Mock(side_effect=RuntimeError("quit failed")),
    )
    initialize, uninitialize = Mock(), Mock()
    monkeypatch.setattr(owned_session.pythoncom, "CoInitializeEx", initialize)
    monkeypatch.setattr(owned_session.pythoncom, "CoUninitialize", uninitialize)
    monkeypatch.setattr(owned_session.win32com.client, "DispatchEx", Mock(return_value=app))
    expected = ValueError if primary_failure else RuntimeError
    message = "original failure" if primary_failure else "close failed"
    with pytest.raises(expected, match=message):
        with owned_session.owned_workbook(Path("work.xlsx")):
            if primary_failure:
                raise ValueError("original failure")
    book.Close.assert_called_once_with(SaveChanges=False)
    app.Quit.assert_called_once()
    uninitialize.assert_called_once()
