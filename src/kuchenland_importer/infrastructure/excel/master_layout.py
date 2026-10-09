"""Use the same header parser with explicit master metadata and payment-role normalization."""

from kuchenland_importer.application.table_schema import ColumnRegistry, find_header, header_key
from kuchenland_importer.application.target_fields import field_key
from kuchenland_importer.domain.master import SheetLayout
from kuchenland_importer.domain.workbook import SourceSheet


def master_registry(registry: ColumnRegistry) -> ColumnRegistry:
    aliases = dict(registry.aliases)
    aliases.update(
        {
            header_key(k): v
            for k, v in {
                "производитель": "mail_supplier",
                "дата": "mail_received",
                "тема письма": "mail_subject",
            }.items()
        }
    )
    return ColumnRegistry(aliases)


def layout_for(sheet: SourceSheet, registry: ColumnRegistry) -> SheetLayout:
    header = find_header(sheet, master_registry(registry))
    if header is None:
        raise ValueError(f"Не найдены заголовки общей книги: {sheet.name}")
    keys = tuple(field_key(key) for key in header.columns)
    if len(set(keys)) != len(keys):
        raise ValueError(f"Неоднозначные поля общей книги: {sheet.name}")
    # Actual calculation templates have service/summary rows 8–10 below the two-row header.
    first = 11 if header.row == 6 and header.depth == 2 else header.row + header.depth
    return SheetLayout(
        sheet.name,
        header.row,
        header.depth,
        first,
        {key: index for index, key in enumerate(keys, 1)},
        len(sheet.rows),
    )
