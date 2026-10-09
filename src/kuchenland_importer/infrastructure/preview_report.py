"""Serialize typed values explicitly, preserving Excel error types."""

import json
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from kuchenland_importer.application.normalize_workbook import NormalizeWorkbook
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.capture_loader import load_attachments
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader


def encode_value(value: object) -> object:
    if isinstance(value, ExcelErrorValue):
        return {"type": "excel_error", "code": value.code}
    if isinstance(value, datetime | date):
        return {"type": "date", "value": value.isoformat()}
    if isinstance(value, Decimal):
        return {"type": "decimal", "value": str(value)}
    raise TypeError(f"Нельзя сериализовать {type(value).__name__}")


@dataclass(frozen=True, slots=True)
class PreviewSummary:
    report: Path
    products: int
    errors: int


def create_preview(manifest: Path, columns: Path, output: Path) -> PreviewSummary:
    try:
        attachments = load_attachments(manifest)
        service = NormalizeWorkbook(load_columns(columns))
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ExcelInputError(f"Не удалось прочитать входные данные: {error}") from error
    result: dict[str, object] = {"schema_version": 1, "stage": "excel_preview"}
    files = []
    product_count = error_count = 0
    for attachment in attachments:
        try:
            preview = service.execute(attachment, WorkbookReader().read(attachment.path))
            products = [
                {
                    "article": p.article,
                    "supplier": p.supplier,
                    "season": p.season,
                    "sheet": p.source_sheet,
                    "row": p.source_row,
                    "values": dict(p.values),
                }
                for p in preview.products
            ]
            issues = [
                {
                    "code": i.code,
                    "message": i.message,
                    "severity": i.severity.value,
                    "source": i.source,
                }
                for i in preview.issues
            ]
            errors = sum(i.severity.value == "error" for i in preview.issues)
            files.append(
                {
                    "attachment": str(attachment.path),
                    "original_name": attachment.original_name,
                    "sha256": attachment.sha256,
                    "mail_subject": attachment.mail.subject,
                    "received_at": attachment.mail.received_at.isoformat(),
                    "products": products,
                    "issues": issues,
                    "eligible_for_import": errors == 0,
                }
            )
            product_count += len(products)
            error_count += errors
        except Exception as error:
            files.append(
                {
                    "attachment": str(attachment.path),
                    "products": [],
                    "eligible_for_import": False,
                    "issues": [
                        {"code": "WORKBOOK_READ_FAILED", "message": str(error), "severity": "error"}
                    ],
                }
            )
            error_count += 1
    result["files"] = files
    result["product_count"] = product_count
    result["error_count"] = error_count
    partial = output.with_suffix(".json.part")
    try:
        with partial.open("x", encoding="utf-8") as stream:
            json.dump(
                result, stream, default=encode_value, ensure_ascii=False, indent=2, allow_nan=False
            )
        partial.rename(output)
    except (OSError, ValueError, TypeError) as error:
        raise ExcelInputError(f"Не удалось сохранить отчёт: {error}") from error
    return PreviewSummary(output, product_count, error_count)
