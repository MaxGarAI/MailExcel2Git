"""Serialize typed values explicitly, preserving Excel error types."""

import json
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from kuchenland_importer.application.bind_photos import BindPhotos
from kuchenland_importer.application.normalize_workbook import NormalizeWorkbook
from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.application.route_workbook import RouteWorkbook
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.infrastructure.capture_loader import load_attachments
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader
from kuchenland_importer.infrastructure.photo_store import PhotoStore


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
    routes_assigned: int = 0
    eligible_products: int = 0
    photos: int = 0


def create_preview(
    manifest: Path, columns: Path, output: Path, *, season_resolver: ResolveSeason | None = None,
    photo_directory: Path | None = None,
) -> PreviewSummary:
    try:
        attachments = load_attachments(manifest)
        registry = load_columns(columns)
        service = NormalizeWorkbook(registry)
        binder = BindPhotos(registry, PhotoStore(photo_directory)) if photo_directory else None
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ExcelInputError(f"Не удалось прочитать входные данные: {error}") from error
    result: dict[str, object] = {
        "schema_version": 1,
        "stage": "photo_preview" if binder is not None
        else "season_preview" if season_resolver is not None else "excel_preview",
    }
    files = []
    product_count = error_count = 0
    routes_assigned = eligible_products = skipped_attachments = 0
    photo_count = eligible_photos = 0
    for attachment in attachments:
        try:
            sheets = WorkbookReader().read(attachment.path)
            preview = service.execute(attachment, sheets)
            if binder is not None:
                preview = binder.execute(
                    preview, sheets, PhotoReader().read(attachment.path, photo_directory)
                )
            all_issues = preview.issues
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
            file_photos = sum(p.photo is not None for p in preview.products)
            if binder is not None:
                for serialized, product in zip(products, preview.products, strict=True):
                    asset = product.photo
                    serialized["photo"] = (
                        {"path": str(asset.path), "pixel_sha256": asset.pixel_sha256,
                         "width": asset.width, "height": asset.height,
                         "image_count": asset.image_count}
                        if asset is not None else None
                    )
            if season_resolver is not None:
                routed = RouteWorkbook(season_resolver).execute(preview)
                all_issues = routed.issues
                for serialized, item in zip(products, routed.products, strict=True):
                    assignment = item.assignment
                    serialized["routing"] = (
                        {
                            "source_season": assignment.source_season,
                            "source_year": assignment.source_year,
                            "effective_year": assignment.effective_year,
                            "effective_season": assignment.effective_season,
                            "route_id": assignment.route.id,
                            "calculation_sheet": assignment.route.calculation_sheet,
                            "photo_sheet": assignment.route.photo_sheet,
                            "sales_start_month": assignment.sales_start_month,
                            "rule": assignment.rule,
                        }
                        if assignment is not None
                        else None
                    )
                routes_assigned += sum(p.assignment is not None for p in routed.products)
            issues = [
                {
                    "code": i.code,
                    "message": i.message,
                    "severity": i.severity.value,
                    "source": i.source,
                }
                for i in all_issues
            ]
            errors = sum(i.severity.value == "error" for i in all_issues)
            if errors:
                skipped_attachments += 1
            else:
                eligible_products += len(products)
                eligible_photos += file_photos
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
            photo_count += file_photos
            error_count += errors
        except Exception as error:
            skipped_attachments += 1
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
    if season_resolver is not None:
        result["routes_assigned"] = routes_assigned
        result["eligible_product_count"] = eligible_products
        result["skipped_attachment_count"] = skipped_attachments
    if binder is not None:
        result["photo_count"] = photo_count
        result["eligible_photo_count"] = eligible_photos
    partial = output.with_suffix(".json.part")
    try:
        with partial.open("x", encoding="utf-8") as stream:
            json.dump(
                result, stream, default=encode_value, ensure_ascii=False, indent=2, allow_nan=False
            )
        partial.rename(output)
    except (OSError, ValueError, TypeError) as error:
        raise ExcelInputError(f"Не удалось сохранить отчёт: {error}") from error
    return PreviewSummary(
        output, product_count, error_count, routes_assigned, eligible_products, photo_count
    )
