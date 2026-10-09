"""Extend existing AutoFilter ranges and reapply the user's supported criteria."""

from typing import Any

from kuchenland_importer.domain.master import MasterSnapshot


def capture_filters(
    book: Any, master: MasterSnapshot
) -> dict[str, tuple[int, int, int, list[dict[str, Any]]]]:
    result = {}
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
            area.AutoFilter(**condition)
