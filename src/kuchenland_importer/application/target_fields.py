"""Build complete target-role fields; photo-table values take precedence on photo tabs."""

import re

from kuchenland_importer.domain.cell_values import CellValue
from kuchenland_importer.domain.season import RoutedProduct


def field_key(key: str) -> str:
    if key == "source:comments#1":
        return "comments"
    if key == "source:quantity#1":
        return "quantity"
    if key == "source:итого / заказ, итого сумма":
        return "source:заказ, итого сумма"
    if re.fullmatch(r"source:заказ сумма, .*", key):
        return "order_amount"
    if re.fullmatch(r"source:[12] платеж / 0 [0-9]+", key):
        return key.split(" / ")[0] + " / amount"
    return key


def product_fields(item: RoutedProduct, photo: bool) -> dict[str, CellValue]:
    product = item.product
    assert item.assignment is not None
    values = {field_key(k): v for k, v in product.values.items() if not k.startswith("photo:")}
    if photo:
        # Calculation-only columns belong on calculation tabs; do not duplicate their entire schema.
        values = {k: v for k, v in values.items() if not k.startswith("source:")}
        values.update(
            {field_key(k[6:]): v for k, v in product.values.items() if k.startswith("photo:")}
        )
    values.pop("photo", None)  # An absent picture must never erase the existing asset.
    values.update(
        {
            "article": product.article,
            "supplier": product.supplier,
            "season": item.assignment.effective_season,
            "mail_supplier": product.supplier,
            "mail_subject": product.attachment.mail.subject,
            "mail_received": product.attachment.mail.received_at.astimezone(),
        }
    )
    return values
