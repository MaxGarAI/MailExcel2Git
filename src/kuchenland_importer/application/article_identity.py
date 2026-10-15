"""Normalize article identity without losing case, punctuation or leading zeroes."""

import re

from kuchenland_importer.domain.workbook import SourceCell


def article_value(cell: SourceCell) -> str:
    value = cell.value
    if isinstance(value, str):
        if not value.strip():
            raise ValueError("Пустой артикул.")
        return value.strip()
    if type(value) is int or isinstance(value, float) and value.is_integer():
        pattern = cell.number_format
        if re.fullmatch(r"0+", pattern):
            return str(int(value)).zfill(len(pattern))
        return str(int(value))
    raise ValueError("Артикул должен быть текстом или целым числом.")

