import json
from datetime import UTC, datetime
from hashlib import sha256
from pathlib import Path

from openpyxl import Workbook

from kuchenland_importer.presentation.cli import main

CONFIG = Path(__file__).resolve().parents[2] / "config" / "default.toml"


def test_cli_routes_each_row_and_skips_entire_invalid_attachment(tmp_path: Path) -> None:
    rows = [
        [("A", "Зима 2026", "10.12.2026"), ("B", "Осень 2027", "")],
        [("C", "Весна", ""), ("D", "ВСЕСЕЗОННЫЙ", "")],
        [("E", "Зима", "")],
    ]
    attachments = []
    originals = {}
    for index, products in enumerate(rows, 1):
        path = tmp_path / f"source-{index}.xlsx"
        book = Workbook()
        book.active.append(["SKU", "Supplier", "Season", "Sales Start"])
        for article, season, start in products:
            book.active.append([article, "Vendor", season, start])
        book.save(path)
        originals[path] = path.read_bytes()
        attachments.append({
            "mail_index": 0, "attachment_index": index, "original_name": path.name,
            "path": path.name, "sha256": sha256(originals[path]).hexdigest(),
        })
    source = tmp_path / "manifest.json"
    source.write_text(json.dumps({
        "schema_version": 1, "stage": "mail_capture",
        "mails": [{
            "entry_id": "entry", "store_id": "store", "subject": "Subject",
            "sender_name": "", "sender_address": "",
            "received_at": datetime(2026, 10, 9, tzinfo=UTC).isoformat(),
        }], "attachments": attachments,
    }), encoding="utf-8")
    source_before = source.read_bytes()
    runtime = tmp_path / "runtime"
    assert main([
        "preview-seasons", "--config", str(CONFIG), "--manifest", str(source),
        "--data-dir", str(runtime),
    ]) == 1
    report = next((runtime / "reports").glob("season-preview-*.json"))
    data = json.loads(report.read_text(encoding="utf-8"))
    assert data["stage"] == "season_preview"
    assert data["product_count"] == 5 and data["routes_assigned"] == 4
    assert data["error_count"] == 1 and data["skipped_attachment_count"] == 1
    assert data["eligible_product_count"] == 3  # C is valid but its attachment is skipped.
    first, second, third = data["files"]
    assert first["eligible_for_import"] and third["eligible_for_import"]
    assert not second["eligible_for_import"]
    assert first["products"][0]["routing"]["route_id"] == "spring"
    assert first["products"][0]["season"] == "Зима 2026"
    assert first["products"][0]["routing"]["effective_season"] == "Весна 2027"
    assert first["products"][0]["routing"]["source_year"] == 2026
    assert first["products"][1]["routing"]["route_id"] == "autumn"
    assert second["products"][1]["routing"] is None
    assert third["issues"][0]["code"] == "WINTER_SALES_START_UNKNOWN"
    assert all(path.read_bytes() == before for path, before in originals.items())
    assert source.read_bytes() == source_before
