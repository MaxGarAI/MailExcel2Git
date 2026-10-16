from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from kuchenland_importer.domain.master import SheetLayout
from kuchenland_importer.infrastructure.excel.row_format import copy_row_format


@pytest.mark.parametrize("fail", [False, True])
def test_new_row_format_copy_uses_destination_and_restores_object_copy_setting(fail):
    source = SimpleNamespace(Copy=Mock(), Address="$A$2:$C$2")
    target = SimpleNamespace(ClearContents=Mock(), Address="$A$3:$C$3")
    if fail:
        source.Copy.side_effect = RuntimeError("Excel copy failed")
    application = SimpleNamespace(CopyObjectsWithCells=True)
    height = SimpleNamespace(RowHeight=80)
    areas = iter((source, target))
    sheet = SimpleNamespace(
        Application=application,
        Name="Products",
        Cells=lambda row, col: (row, col),
        Range=lambda *args: next(areas),
        Rows=lambda row: height,
    )
    layout = SheetLayout("Products", 1, 1, 2, {"article": 1, "name": 3}, 2)
    if fail:
        with pytest.raises(ValueError, match="Products.*A"):
            copy_row_format(sheet, 2, 3, layout)
        target.ClearContents.assert_not_called()
    else:
        copy_row_format(sheet, 2, 3, layout)
        target.ClearContents.assert_called_once()
    source.Copy.assert_called_once_with(Destination=target)
    assert application.CopyObjectsWithCells is True
