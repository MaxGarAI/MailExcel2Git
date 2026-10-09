"""Embed receipt provenance in the same atomic workbook transaction."""

from datetime import UTC
from typing import Any

from kuchenland_importer.domain.master import MasterSnapshot
from kuchenland_importer.domain.write_plan import WritePlan
from kuchenland_importer.infrastructure.excel.com_values import write_value
from kuchenland_importer.infrastructure.excel.master_snapshot import STATE_SHEET


def save_state(book: Any, master: MasterSnapshot, plan: WritePlan) -> None:
    try:
        sheet = book.Worksheets(STATE_SHEET)
    except Exception:
        sheet = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
        sheet.Name = STATE_SHEET
    entries = {
        article: (p.received_at.isoformat(), p.fingerprint)
        for article, p in master.products.items()
        if p.received_at is not None
    }
    for action in plan.products:
        mail = action.incoming.product.attachment.mail
        entries[action.incoming.product.article] = (
            mail.received_at.astimezone(UTC).isoformat(),
            action.fingerprint,
        )
    sheet.Cells.ClearContents()
    sheet.Cells(1, 1).Value2 = "kuchenland-import-state-v1"
    for row, (article, (received, fingerprint)) in enumerate(sorted(entries.items()), 2):
        for column, value in enumerate((article, received, fingerprint), 1):
            write_value(sheet.Cells(row, column), value)
    sheet.Visible = 2  # xlSheetVeryHidden; not a supplier/user product table.
