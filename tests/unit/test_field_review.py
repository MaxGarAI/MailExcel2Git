from dataclasses import replace

import pytest
from test_import_planning import master, prepared

from kuchenland_importer.application.review_fields import review_fields
from kuchenland_importer.application.table_schema import ColumnRegistry
from kuchenland_importer.application.target_fields import product_fields
from kuchenland_importer.domain.field_review import FieldDecision, rule_key
from kuchenland_importer.infrastructure.field_rules import load_rules, save_rules


def batch_with_extra():
    item = prepared()
    routed = item.products[0]
    product = replace(routed.product, values={**routed.product.values, "source:цвет": "зелёный"})
    return (replace(item, products=(replace(routed, product=product),)),)


def decide(choice):
    if choice.source == "source:цвет":
        return FieldDecision("create", "цвет")
    return FieldDecision("skip")


def test_create_reopen_repeat_does_not_duplicate_column_or_change_values(tmp_path):
    snapshot, rules = master(), {}
    result = review_fields(batch_with_extra(), snapshot, tmp_path / "master.xlsx", rules,
                           decide, ColumnRegistry({}))
    assert len(result.columns) == 1
    assert result.columns[0].sheet == "Весна"
    assert product_fields(result.batch[0].products[0], False)["source:цвет"] == "зелёный"
    assert "source:цвет" not in product_fields(result.batch[0].products[0], True)
    path = tmp_path / "rules.json"
    save_rules(path, rules)
    restored = load_rules(path)
    assert restored == rules
    def unexpected(choice):
        pytest.fail("Persisted decision must not prompt")
    repeat = review_fields(batch_with_extra(), snapshot, tmp_path / "master.xlsx", restored,
                           unexpected, ColumnRegistry({}))
    assert not repeat.columns
    assert repeat.batch[0].products[0].product.values == result.batch[0].products[0].product.values


def test_mapping_supplier_scope_and_collision(tmp_path):
    snapshot = master()
    snapshot.layouts["Весна"].columns["source:оттенок"] = 3
    def mapping(choice):
        if choice.source == "source:цвет":
            return FieldDecision("map", "source:оттенок")
        return FieldDecision("skip")
    rules = {}
    result = review_fields(batch_with_extra(), snapshot, tmp_path / "m.xlsx", rules,
                           mapping, ColumnRegistry({}))
    fields = product_fields(result.batch[0].products[0], False)
    assert fields["source:оттенок"] == "зелёный" and "source:цвет" not in fields
    assert not result.columns
    assert rule_key(tmp_path / "m.xlsx", "Other", False, "source:цвет") not in rules
    assert rule_key(tmp_path / "other.xlsx", "Vendor", False, "source:цвет") not in rules


@pytest.mark.parametrize("title", ["", "=FORMULA", "SKU", "Группа / Подгруппа"])
def test_invalid_new_header_is_rejected(tmp_path, title):
    def decision(choice):
        return FieldDecision("create", title)
    with pytest.raises(ValueError):
        review_fields(batch_with_extra(), master(), tmp_path / "m.xlsx", {}, decision,
                      ColumnRegistry({"sku": "article"}))


def test_damaged_rules_not_silently_overwritten(tmp_path):
    path = tmp_path / "rules.json"
    path.write_text('{"schema_version":99}', encoding="utf-8")
    with pytest.raises(ValueError):
        load_rules(path)
