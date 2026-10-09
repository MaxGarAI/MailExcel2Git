"""Read a sales month/year without treating bare numbers as Excel serial dates."""

import re
from datetime import date, datetime

from kuchenland_importer.domain.cell_values import CellValue
from kuchenland_importer.domain.season import normalize_alias

MONTHS = (
    ("январь", "января", "january"),
    ("февраль", "февраля", "february"),
    ("март", "марта", "march"),
    ("апрель", "апреля", "april"),
    ("май", "мая", "may"),
    ("июнь", "июня", "june"),
    ("июль", "июля", "july"),
    ("август", "августа", "august"),
    ("сентябрь", "сентября", "september"),
    ("октябрь", "октября", "october"),
    ("ноябрь", "ноября", "november"),
    ("декабрь", "декабря", "december"),
)


def sales_period(value: CellValue) -> tuple[int, int | None] | None:
    if isinstance(value, date):
        return value.month, value.year
    if not isinstance(value, str) or not value.strip():
        return None
    text = normalize_alias(value)
    for pattern in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%m.%Y", "%Y-%m"):
        try:
            parsed = datetime.strptime(text, pattern)
            return parsed.month, parsed.year
        except ValueError:
            continue
    # A day/month value has no season year; 2000 permits a leap-day value.
    if re.fullmatch(r"\d{1,2}\.\d{1,2}", text):
        try:
            return datetime.strptime(text + ".2000", "%d.%m.%Y").month, None
        except ValueError:
            return None
    match = re.fullmatch(r"([a-zа-я]+)(?:\s+([1-9]\d{3}))?", text)
    if match is not None:
        for month, names in enumerate(MONTHS, 1):
            if match[1] in names:
                return month, int(match[2]) if match[2] else None
    return None


def sales_month(value: CellValue) -> int | None:
    period = sales_period(value)
    return period[0] if period is not None else None
