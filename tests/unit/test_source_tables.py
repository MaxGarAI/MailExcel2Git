from pathlib import Path

from kuchenland_importer.application.source_tables import classify_source_tables
from kuchenland_importer.domain.workbook import SourceCell, SourceSheet
from kuchenland_importer.infrastructure.excel.columns import load_columns


def sheet(name, rows):
    return SourceSheet(name, tuple(tuple(SourceCell(value) for value in row) for row in rows), ())


def comparison(article="OLD"):
    return sheet(
        "Reference",
        [
            ["Article", "Photo", "Name", "Article", "Photo", "Name"],
            [article, None, "Old", None, None, None],
        ],
    )


def primary():
    return sheet("Order", [["Article", "Supplier", "Season"], ["A", "Vendor", "Весна"]])


def test_unrelated_comparison_is_warning_and_does_not_taint_identity():
    result = classify_source_tables(
        (comparison(), primary()), load_columns(Path("config/columns.toml"))
    )
    assert set(result.headers) == {"Order"}
    assert result.articles == {"A"} and result.identity_complete
    assert [i.code for i in result.issues] == ["REFERENCE_SHEET_SKIPPED"]
    assert result.issues[0].severity.value == "warning"


def test_comparison_containing_an_imported_article_remains_an_explicit_conflict():
    result = classify_source_tables(
        (primary(), comparison("A")), load_columns(Path("config/columns.toml"))
    )
    assert "SHEET_SCHEMA" in {i.code for i in result.issues}
    assert not result.identity_complete


def test_ambiguous_primary_is_not_treated_as_reference():
    ambiguous = sheet(
        "Bad",
        [
            ["Article", "SKU", "Supplier", "Season", "Photo", "Photo", "Name", "Name"],
            ["X", "Y", "Vendor", "Весна", None, None, "One", "Two"],
        ],
    )
    result = classify_source_tables(
        (primary(), ambiguous), load_columns(Path("config/columns.toml"))
    )
    assert any(i.code == "SHEET_SCHEMA" and i.source == "Bad" for i in result.issues)


def test_comparison_without_primary_identity_cannot_be_silently_ignored():
    result = classify_source_tables((comparison(),), load_columns(Path("config/columns.toml")))
    assert not result.headers
    assert [i.code for i in result.issues] == ["SHEET_SCHEMA"]


def test_comparison_with_uncached_article_cannot_be_proved_unrelated():
    helper = comparison()
    rows = list(helper.rows)
    rows[1] = (SourceCell(None, missing_formula_cache=True), *rows[1][1:])
    helper = SourceSheet(helper.name, tuple(rows), ())
    result = classify_source_tables((primary(), helper), load_columns(Path("config/columns.toml")))
    assert "REFERENCE_SHEET_SKIPPED" not in {i.code for i in result.issues}
    assert "SHEET_SCHEMA" in {i.code for i in result.issues}
