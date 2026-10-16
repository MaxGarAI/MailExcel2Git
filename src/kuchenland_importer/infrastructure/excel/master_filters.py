"""Extend existing AutoFilter ranges and reapply the user's supported criteria."""

from pathlib import Path
from typing import Any

import pythoncom
from win32com.client import VARIANT

from kuchenland_importer.domain.master import MasterSnapshot
from kuchenland_importer.infrastructure.excel.filter_metadata import read_top_filters

# IDispatch optional argument, unlike VT_EMPTY or the InvokeTypes-only PyOleMissing.
MISSING = VARIANT(pythoncom.VT_ERROR, -2147352572)  # DISP_E_PARAMNOTFOUND


def capture_filters(
    book: Any, master: MasterSnapshot
) -> dict[str, tuple[int, int, int, list[dict[str, Any]]]]:
    result = {}
    top_filters = None
    for name in master.layouts:
        sheet = book.Worksheets(name)
        if not sheet.AutoFilterMode:
            continue
        auto = sheet.AutoFilter
        area = auto.Range
        filters = []
        for field in range(1, auto.Filters.Count + 1):
            item = auto.Filters(field)
            if not item.On:
                continue
            criteria: dict[str, Any] = {"Field": field, "Operator": item.Operator}
            for key in ("Criteria1", "Criteria2"):
                try:
                    criteria[key] = getattr(item, key)
                except Exception:
                    continue  # Excel omits the unused side of a one-sided filter.
            if criteria["Operator"] in {8, 9} and "Criteria1" in criteria:
                # Excel returns Interior/Font objects, not the input RGB integer.
                # Resolve Color while the original filter COM object is still valid.
                color = criteria["Criteria1"]
                criteria["Criteria1"] = int(color if isinstance(color, int) else color.Color)
            if criteria["Operator"] in {3, 4, 5, 6}:
                # COM reports the computed threshold (e.g. >=70), not the requested
                # number/percentage (e.g. top 3). Recover the original saved parameter.
                if top_filters is None:
                    top_filters = read_top_filters(Path(str(book.FullName)))
                saved = top_filters.get((name, field))
                if saved is None or saved[0] != criteria["Operator"]:
                    raise ValueError(f"Нет сохранённых параметров Top/Bottom фильтра {name}, "
                                     f"поле {field}.")
                criteria["Criteria1"] = saved[1]
                criteria.pop("Criteria2", None)
            if not {"Criteria1", "Criteria2"} & criteria.keys():
                raise ValueError(f"Невозможно сохранить критерий фильтра {name}, поле {field}.")
            filters.append(criteria)
        result[name] = (int(area.Row), int(area.Column), int(area.Columns.Count), filters)
    return result


def restore_filters(
    book: Any,
    master: MasterSnapshot,
    filters: dict[str, tuple[int, int, int, list[dict[str, Any]]]],
) -> None:
    for name, (header, column, width, criteria) in filters.items():
        sheet = book.Worksheets(name)
        layout = master.layouts[name]
        last = max(
            header + 1, int(sheet.Cells(sheet.Rows.Count, layout.columns["article"]).End(-4162).Row)
        )
        sheet.AutoFilterMode = False
        area = sheet.Range(sheet.Cells(header, column), sheet.Cells(last, column + width - 1))
        sheet.Activate()
        # Invoke with truly omitted arguments; generated wrappers inject invalid defaults
        # for this optional-argument method on some localized Office builds.
        dispatch_id = area._oleobj_.GetIDsOfNames("AutoFilter")
        try:
            area._oleobj_.Invoke(dispatch_id, 0, 1, True)
        except Exception as error:
            raise ValueError(f"Не удалось восстановить фильтр {name}: {area.Address}") from error
        for condition in criteria:
            apply_criteria(area, condition, sheet_name=name)


def apply_criteria(area: Any, condition: dict[str, Any], *, sheet_name: str) -> None:
    # Filter.Operator returns zero for one criterion. Zero is not a valid input
    # XlAutoFilterOperator: omit it rather than replaying it or injecting xlAnd.
    arguments = [
        condition["Field"],
        condition.get("Criteria1", MISSING),
        condition.get("Operator") or MISSING,
        condition.get("Criteria2", MISSING),
    ]
    while arguments and arguments[-1] is MISSING:
        arguments.pop()
    try:
        dispatch_id = area._oleobj_.GetIDsOfNames("AutoFilter")
        area._oleobj_.Invoke(dispatch_id, 0, 1, True, *arguments)
    except Exception as error:
        raise ValueError(
            f"Не удалось восстановить критерий фильтра {sheet_name}: "
            f"поле {condition['Field']}, диапазон {area.Address}"
        ) from error


def suspend_filters(book: Any, filters: dict[str, Any]) -> None:
    """Expose filtered rows before End(xlUp), row replacement and picture insertion."""
    for name in filters:
        sheet = book.Worksheets(name)
        if sheet.FilterMode:
            try:
                sheet.ShowAllData()
            except Exception as error:
                raise ValueError(f"Не удалось временно снять фильтрацию листа {name}.") from error
