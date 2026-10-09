from dataclasses import replace
from datetime import UTC, datetime, timedelta
from pathlib import Path

from kuchenland_importer.application.plan_import import make_plan
from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.application.select_latest import SelectLatest
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.import_batch import PreparedAttachment
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.master import (
    ExistingProduct,
    ExistingRow,
    MasterSnapshot,
    SheetLayout,
)
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.season import RoutedProduct, SeasonRoute

ROUTE = SeasonRoute("spring", ("Весна",), "Весна", "ФОТО весна")


def prepared(*, day=1, price=10, eligible=True, article="001-A", unknown=False):
    mail = MailMetadata(str(day), "store", "Subject", datetime(2026, 1, day, tzinfo=UTC), "", "")
    attachment = SavedAttachment(mail, 1, "source.xlsx", Path(__file__).resolve(), "0" * 64)
    product = ProductRecord(
        article,
        "Vendor",
        "Весна",
        attachment,
        "Товары",
        2,
        {"article": article, "season": "Весна", "purchase_price": price},
    )
    routed = RoutedProduct(product, ResolveSeason((ROUTE,)).execute(product))
    issues = () if eligible else (ImportIssue("BAD", "Ошибка", Severity.ERROR, "source"),)
    return PreparedAttachment(
        attachment,
        (routed,) if not unknown else (),
        frozenset({article}) if not unknown else frozenset(),
        issues,
    )


def master():
    return MasterSnapshot(
        {
            "Весна": SheetLayout("Весна", 1, 1, 2, {"article": 1, "season": 2}, 3),
            "ФОТО весна": SheetLayout("ФОТО весна", 1, 1, 2, {"article": 1, "photo": 2}, 3),
        },
        {},
        ("Весна", "ФОТО весна"),
    )


def test_latest_mail_wins_regardless_of_input_order():
    old, new = prepared(day=1), prepared(day=2, price=20)
    for files in ((old, new), (new, old)):
        result = SelectLatest().execute(files)
        assert result.products[0].product.values["purchase_price"] == 20


def test_unsupported_mapped_error_rejects_only_its_attachment():
    item = prepared()
    routed = item.products[0]
    product = replace(
        routed.product,
        values={**routed.product.values, "purchase_price": ExcelErrorValue("#SPILL!")},
    )
    bad = replace(item, products=(replace(routed, product=product),))
    good = prepared(article="OTHER")
    snapshot = master()
    snapshot.layouts["Весна"].columns["purchase_price"] = 3
    plan = make_plan((bad, good), snapshot, frozenset({"#DIV/0!"}))
    assert [p.incoming.product.article for p in plan.products] == ["OTHER"]
    assert "EXCEL_ERROR_UNSUPPORTED" in {i.code for i in plan.issues}


def test_unsupported_error_in_skipped_field_does_not_reject_attachment():
    item = prepared()
    routed = item.products[0]
    product = replace(
        routed.product,
        values={**routed.product.values, "purchase_price": ExcelErrorValue("#SPILL!")},
    )
    plan = make_plan(
        (replace(item, products=(replace(routed, product=product),)),),
        master(),
        frozenset({"#DIV/0!"}),
    )
    assert len(plan.products) == 1


def test_rejected_newest_blocks_older_version():
    result = SelectLatest().execute((prepared(), prepared(day=2, eligible=False)))
    assert not result.products
    assert result.skipped_articles == 1
    assert "NEWEST_ATTACHMENT_REJECTED" in {i.code for i in result.issues}


def test_unreadable_newest_does_not_allow_unknown_fallback():
    result = SelectLatest().execute((prepared(), prepared(day=2, eligible=False, unknown=True)))
    assert not result.products


def test_newer_valid_mail_can_pass_unreadable_older_mail():
    result = SelectLatest().execute((prepared(day=2), prepared(eligible=False, unknown=True)))
    assert len(result.products) == 1


def test_partial_article_identity_does_not_allow_unknown_fallback():
    bad = replace(prepared(day=2, article="B", eligible=False), identity_complete=False)
    result = SelectLatest().execute((prepared(), bad))
    assert not result.products


def test_historical_master_date_prevents_older_mail_overwrite():
    snapshot = master()
    from datetime import date

    snapshot.products["001-A"] = ExistingProduct(
        [ExistingRow("Весна", 2), ExistingRow("ФОТО весна", 2)],
        received_day=date(2026, 1, 3),
    )
    assert make_plan((prepared(day=1),), snapshot).unchanged == 1


def test_equal_timestamp_conflicting_values_are_rejected():
    result = SelectLatest().execute((prepared(price=1), prepared(price=2)))
    assert not result.products
    assert result.issues[0].code == "MAIL_TIME_CONFLICT"


def test_equal_timestamp_identical_products_are_one_product():
    assert len(SelectLatest().execute((prepared(), prepared())).products) == 1


def test_missing_pair_rejects_entire_attachment():
    snapshot = master()
    del snapshot.layouts["ФОТО весна"]
    result = make_plan((prepared(),), snapshot)
    assert not result.products
    assert "MASTER_SHEET_MISSING" in {i.code for i in result.issues}


def test_valid_attachment_survives_another_invalid_file():
    result = make_plan((prepared(), prepared(article="B", eligible=False)), master())
    assert [p.incoming.product.article for p in result.products] == ["001-A"]


def test_existing_mirrors_are_one_product_and_duplicates_are_errors():
    snapshot = master()
    old = ExistingProduct([ExistingRow("Весна", 2), ExistingRow("ФОТО весна", 2)])
    snapshot.products["001-A"] = old
    assert len(make_plan((prepared(),), snapshot).products) == 1
    old.rows.append(ExistingRow("Весна", 3))
    assert not make_plan((prepared(),), snapshot).products


def test_single_existing_photo_row_is_repaired_as_one_updated_product():
    snapshot = master()
    snapshot.products["001-A"] = ExistingProduct([ExistingRow("ФОТО весна", 2)])
    result = make_plan((prepared(),), snapshot)
    assert len(result.products) == 1 and result.products[0].previous is not None
    assert "MASTER_MIRROR_REPAIRED" in {i.code for i in result.issues}


def test_saved_receipt_state_makes_repeated_import_idempotent():
    snapshot = master()
    item = prepared()
    first = make_plan((item,), snapshot)
    old = ExistingProduct(
        [ExistingRow("Весна", 2), ExistingRow("ФОТО весна", 2)],
        received_at=item.attachment.mail.received_at,
        fingerprint=first.products[0].fingerprint,
    )
    snapshot.products["001-A"] = old
    second = make_plan((item,), snapshot)
    assert not second.products and second.unchanged == 1
    older = replace(
        item,
        attachment=replace(
            item.attachment,
            mail=replace(
                item.attachment.mail,
                received_at=item.attachment.mail.received_at - timedelta(days=1),
            ),
        ),
    )
    assert make_plan((older,), snapshot).unchanged == 1


def test_changed_data_at_saved_timestamp_is_rejected():
    snapshot = master()
    item = prepared()
    first = make_plan((item,), snapshot)
    snapshot.products["001-A"] = ExistingProduct(
        [ExistingRow("Весна", 2), ExistingRow("ФОТО весна", 2)],
        received_at=item.attachment.mail.received_at,
        fingerprint=first.products[0].fingerprint,
    )
    result = make_plan((prepared(price=99),), snapshot)
    assert not result.products
    assert "MASTER_TIME_CONFLICT" in {i.code for i in result.issues}


def test_article_case_and_leading_zeros_do_not_collapse():
    batch = (
        prepared(article="001"),
        prepared(article="1"),
        prepared(article="Ab"),
        prepared(article="ab"),
    )
    assert len(SelectLatest().execute(batch).products) == 4
