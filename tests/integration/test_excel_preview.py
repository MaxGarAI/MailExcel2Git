import json
from datetime import UTC, datetime
from hashlib import sha256
from pathlib import Path

import pytest
from openpyxl import Workbook

from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.preview_report import create_preview

CONFIG = Path(__file__).resolve().parents[2] / "config" / "columns.toml"


def manifest(tmp_path: Path) -> Path:
    good = tmp_path / "good.xlsx"
    book = Workbook()
    book.active.append(["Артикул", "Поставщик", "Сезон"])
    book.active.append(["A", "Vendor", "Весна"])
    book.save(good)
    bad = tmp_path / "broken.xlsx"
    bad.write_bytes(b"not Excel")
    path = tmp_path / "manifest.json"
    path.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "stage": "mail_capture",
                "mails": [
                    {
                        "entry_id": "entry",
                        "store_id": "store",
                        "subject": "",
                        "sender_name": "",
                        "sender_address": "",
                        "received_at": datetime.now(UTC).isoformat(),
                    }
                ],
                "attachments": [
                    {
                        "mail_index": 0,
                        "attachment_index": index,
                        "original_name": p.name,
                        "path": p.name,
                        "sha256": sha256(p.read_bytes()).hexdigest(),
                    }
                    for index, p in enumerate((good, bad), 1)
                ],
            }
        ),
        encoding="utf-8",
    )
    return path


def test_bad_workbook_does_not_block_good_attachment(tmp_path: Path) -> None:
    source = manifest(tmp_path)
    result = create_preview(source, CONFIG, tmp_path / "preview.json")
    assert result.products == 1
    assert result.errors == 1
    data = json.loads(result.report.read_text(encoding="utf-8"))
    assert data["files"][0]["eligible_for_import"] is True
    assert data["files"][1]["eligible_for_import"] is False
    assert not list(tmp_path.glob("*.part"))


def test_modified_source_is_detected_before_processing(tmp_path: Path) -> None:
    source = manifest(tmp_path)
    (tmp_path / "good.xlsx").write_bytes(b"changed")
    with pytest.raises(ExcelInputError, match="сумма"):
        create_preview(source, CONFIG, tmp_path / "preview.json")
    assert not (tmp_path / "preview.json").exists()
