from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from kuchenland_importer.infrastructure.excel.master_filters import MISSING, apply_criteria


@pytest.mark.parametrize(
    "condition,expected",
    [
        ({"Field": 2, "Criteria1": "=Red", "Operator": 0}, (2, "=Red")),
        ({"Field": 1, "Criteria1": 0, "Operator": 0}, (1, 0)),
        ({"Field": 2, "Criteria1": "=", "Operator": 0}, (2, "=")),
        ({"Field": 2, "Criteria1": "<>", "Operator": 0}, (2, "<>")),
        (
            {"Field": 3, "Criteria2": (1, "10/31/2026"), "Operator": 0},
            (3, MISSING, MISSING, (1, "10/31/2026")),
        ),
        (
            {"Field": 2, "Criteria1": ("Red", "Blue", "Green"), "Operator": 7},
            (2, ("Red", "Blue", "Green"), 7),
        ),
        ({"Field": 4, "Criteria1": ">20", "Criteria2": "<80", "Operator": 1}, (4, ">20", 1, "<80")),
    ],
)
def test_filter_restoration_omits_only_unused_com_arguments(condition, expected):
    dispatch = Mock()
    dispatch.GetIDsOfNames.return_value = 3289
    area = SimpleNamespace(_oleobj_=dispatch, Address="$A$1:$E$100")
    apply_criteria(area, condition, sheet_name="Products")
    dispatch.Invoke.assert_called_once_with(3289, 0, 1, True, *expected)


def test_failed_filter_restoration_has_context_and_keeps_original_exception():
    dispatch = Mock()
    cause = RuntimeError("COM failure")
    dispatch.Invoke.side_effect = cause
    area = SimpleNamespace(_oleobj_=dispatch, Address="$A$1:$E$100")
    with pytest.raises(ValueError, match="Products.*поле 2.*A") as caught:
        apply_criteria(
            area, {"Field": 2, "Criteria1": "private value", "Operator": 0}, sheet_name="Products"
        )
    assert caught.value.__cause__ is cause
    assert "private value" not in str(caught.value)


@pytest.mark.parametrize("operator", [8, 9])
def test_color_filter_is_captured_before_its_com_object_is_invalidated(operator):
    from kuchenland_importer.infrastructure.excel.master_filters import capture_filters

    item = SimpleNamespace(On=True, Operator=operator, Criteria1=SimpleNamespace(Color=255))
    filters = Mock(Count=1, return_value=item)
    area = SimpleNamespace(Row=1, Column=1, Columns=SimpleNamespace(Count=1))
    sheet = SimpleNamespace(
        AutoFilterMode=True, AutoFilter=SimpleNamespace(Range=area, Filters=filters)
    )
    book = SimpleNamespace(Worksheets=lambda name: sheet)
    snapshot = SimpleNamespace(layouts={"Products": None})
    captured = capture_filters(book, snapshot)
    assert captured["Products"][3][0]["Criteria1"] == 255


@pytest.mark.parametrize(
    "top,percent,operator",
    [
        (True, False, 3),
        (True, True, 4),
        (False, False, 5),
        (False, True, 6),
    ],
)
def test_top_bottom_count_is_read_instead_of_replaying_com_threshold(
    tmp_path, top, percent, operator
):
    from openpyxl import Workbook
    from openpyxl.worksheet.filters import FilterColumn, Top10

    from kuchenland_importer.infrastructure.excel.master_filters import capture_filters

    path = tmp_path / "master.xlsx"
    book = Workbook()
    book.active.title = "Products"
    book.active.append(["Price"])
    book.active.auto_filter.ref = "A1:A10"
    book.active.auto_filter.filterColumn.append(
        FilterColumn(colId=0, top10=Top10(top=top, percent=percent, val=3, filterVal=70))
    )
    book.save(path)
    book.close()
    item = SimpleNamespace(On=True, Operator=operator, Criteria1=">=70")
    filters = Mock(Count=1, return_value=item)
    area = SimpleNamespace(Row=1, Column=1, Columns=SimpleNamespace(Count=1))
    sheet = SimpleNamespace(
        AutoFilterMode=True, AutoFilter=SimpleNamespace(Range=area, Filters=filters)
    )
    excel = SimpleNamespace(Worksheets=lambda name: sheet, FullName=str(path))
    captured = capture_filters(excel, SimpleNamespace(layouts={"Products": None}))
    assert float(captured["Products"][3][0]["Criteria1"]) == 3
    assert captured["Products"][3][0]["Operator"] == operator


@pytest.mark.parametrize("active", [False, True])
def test_filter_is_suspended_only_when_it_is_actually_filtering_rows(active):
    from kuchenland_importer.infrastructure.excel.master_filters import suspend_filters

    sheet = SimpleNamespace(FilterMode=active, ShowAllData=Mock())
    suspend_filters(SimpleNamespace(Worksheets=lambda name: sheet), {"Products": ()})
    assert sheet.ShowAllData.call_count == int(active)


def test_filter_suspension_failure_is_not_ignored():
    from kuchenland_importer.infrastructure.excel.master_filters import suspend_filters

    cause = RuntimeError("COM failure")
    sheet = SimpleNamespace(FilterMode=True, ShowAllData=Mock(side_effect=cause))
    with pytest.raises(ValueError, match="Products") as caught:
        suspend_filters(SimpleNamespace(Worksheets=lambda name: sheet), {"Products": ()})
    assert caught.value.__cause__ is cause
