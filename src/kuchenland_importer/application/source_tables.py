"""Classify primary/photo tables before validating unrelated comparison panels."""

from dataclasses import dataclass

from kuchenland_importer.application.article_identity import article_value
from kuchenland_importer.application.table_schema import ColumnRegistry, TableHeader, find_header
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.workbook import SourceSheet

PRIMARY = frozenset({"article", "supplier", "season"})
PHOTO = frozenset({"article", "photo"})


@dataclass(frozen=True, slots=True)
class SourceTables:
    headers: dict[str, TableHeader]
    articles: frozenset[str]
    identity_complete: bool
    issues: tuple[ImportIssue, ...]


def unrelated_comparison(sheet: SourceSheet, registry: ColumnRegistry, articles: set[str]) -> bool:
    if not articles:
        return False
    for index, row in enumerate(sheet.rows[:50]):
        fields = [registry.resolve(str(c.value)) if c.value is not None else None for c in row]
        if (
            fields.count("article") < 2
            or fields.count("photo") < 2
            or fields.count("name") < 2
            or PRIMARY <= set(fields)
        ):
            continue
        columns = [i for i, name in enumerate(fields) if name == "article"]
        for values in sheet.rows[index + 1 :]:
            for column in columns:
                cell = values[column]
                if cell.missing_formula_cache:
                    return False
                if cell.value in {None, ""}:
                    continue
                try:
                    if article_value(cell) in articles:
                        return False
                except ValueError:
                    return False
        return True
    return False


def classify_source_tables(
    sheets: tuple[SourceSheet, ...], registry: ColumnRegistry
) -> SourceTables:
    headers: dict[str, TableHeader] = {}
    articles: set[str] = set()
    invalid: set[str] = set()
    issues = []
    complete = True
    for sheet in sheets:
        try:
            header = find_header(sheet, registry, required_fields=PRIMARY)
        except ValueError as error:
            invalid.add(sheet.name)
            issues.append(ImportIssue("SHEET_SCHEMA", str(error), Severity.ERROR, sheet.name))
            complete = False
            continue
        if header is None:
            continue
        headers[sheet.name] = header
        for row in sheet.rows[header.row + header.depth - 1 :]:
            cell = row[header.columns.index("article")]
            if cell.value in {None, ""}:
                continue
            try:
                articles.add(article_value(cell))
            except ValueError:
                complete = False
    for sheet in sheets:
        if sheet.name in headers or sheet.name in invalid:
            continue
        if unrelated_comparison(sheet, registry, articles):
            issues.append(
                ImportIssue(
                    "REFERENCE_SHEET_SKIPPED",
                    "Сравнительный лист не содержит артикулов импортируемых товаров и пропущен.",
                    Severity.WARNING,
                    sheet.name,
                )
            )
            continue
        try:
            header = find_header(sheet, registry, required_fields=PHOTO)
            if header is not None:
                headers[sheet.name] = header
        except ValueError as error:
            issues.append(ImportIssue("SHEET_SCHEMA", str(error), Severity.ERROR, sheet.name))
            complete = False
    return SourceTables(headers, frozenset(articles), complete, tuple(issues))
