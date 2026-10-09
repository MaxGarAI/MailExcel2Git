from datetime import UTC, datetime
from decimal import Decimal
from pathlib import Path
from typing import cast

import pytest

from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.product import CellValue, ProductRecord
from kuchenland_importer.domain.report import ImportIssue, ImportReport, Severity


def attachment(tmp_path: Path) -> SavedAttachment:
    mail = MailMetadata("entry", "store", "Тема", datetime(2026, 10, 9, tzinfo=UTC), "Закупщик", "")
    return SavedAttachment(mail, 1, "товары.xlsx", tmp_path / "товары.xlsx", "a" * 64)


def test_article_zeros_and_defensive_copy(tmp_path: Path) -> None:
    values: dict[str, CellValue] = {"price": Decimal("15.20")}
    product = ProductRecord(
        " 001-Ab ", "Поставщик", "Весна 2027", attachment(tmp_path), "Товары", 2, values
    )
    values["price"] = Decimal("999")
    assert product.article == "001-Ab"
    assert product.values["price"] == Decimal("15.20")
    with pytest.raises(TypeError):
        cast(dict[str, CellValue], product.values)["price"] = 0


@pytest.mark.parametrize("value", [float("nan"), float("inf"), Decimal("NaN")])
def test_non_finite_values_rejected(tmp_path: Path, value: CellValue) -> None:
    with pytest.raises(ValueError):
        ProductRecord(
            "001", "Поставщик", "Весна", attachment(tmp_path), "Лист", 2, {"price": value}
        )


def test_received_date_requires_timezone() -> None:
    with pytest.raises(ValueError, match="часовой пояс"):
        MailMetadata("entry", "store", "", datetime(2026, 10, 9), "", "")


def test_report_separates_errors_and_warnings() -> None:
    report = ImportReport(
        "run",
        issues=(
            ImportIssue("NO_PHOTO", "Фото отсутствует", Severity.WARNING, "book.xlsx:2"),
            ImportIssue("NO_SKU", "Нет артикула", Severity.ERROR, "book.xlsx:3"),
        ),
    )
    assert report.error_count == 1
    assert report.warning_count == 1


def test_report_rejects_negative_count() -> None:
    with pytest.raises(ValueError):
        ImportReport("run", products_added=-1)
