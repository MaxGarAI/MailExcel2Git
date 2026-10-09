"""Validate whole attachments against the master before selecting newest candidates."""

from dataclasses import replace

from kuchenland_importer.application.product_fingerprint import fingerprint_json
from kuchenland_importer.application.select_latest import SelectLatest
from kuchenland_importer.application.target_fields import product_fields
from kuchenland_importer.domain.cell_values import ExcelErrorValue
from kuchenland_importer.domain.import_batch import PreparedAttachment
from kuchenland_importer.domain.master import MasterSnapshot
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.write_plan import ProductWrite, WritePlan


def make_plan(
    batch: tuple[PreparedAttachment, ...],
    master: MasterSnapshot,
    supported_errors: frozenset[str] | None = None,
) -> WritePlan:
    validated = []
    for item in batch:
        issues = list(item.issues)
        for routed in item.products:
            product, assignment = routed.product, routed.assignment
            if assignment is None:
                continue
            pair = (assignment.route.calculation_sheet, assignment.route.photo_sheet)
            for name, photo in ((pair[0], False), (pair[1], True)):
                if name in master.layouts:
                    fields = product_fields(routed, photo)
                    unsupported = {
                        value.code
                        for key, value in fields.items()
                        if key in master.layouts[name].columns
                        and isinstance(value, ExcelErrorValue)
                        and supported_errors is not None
                        and value.code not in supported_errors
                    }
                    if unsupported:
                        issues.append(
                            ImportIssue(
                                "EXCEL_ERROR_UNSUPPORTED",
                                f"Неподдерживаемые ошибки Excel: {sorted(unsupported)}",
                                Severity.ERROR,
                                product.article,
                            )
                        )
                    missing = (
                        set(product_fields(routed, photo)) - master.layouts[name].columns.keys()
                    )
                    if missing:
                        issues.append(
                            ImportIssue(
                                "FIELDS_SKIPPED",
                                f"Поля без столбцов пропущены: {sorted(missing)}",
                                Severity.WARNING,
                                f"{product.article}: {name}",
                            )
                        )
            if any(name not in master.layouts for name in pair):
                issues.append(
                    ImportIssue(
                        "MASTER_SHEET_MISSING",
                        "Нет согласованной пары листов сезона.",
                        Severity.ERROR,
                        product.article,
                    )
                )
            old = master.products.get(product.article)
            if old is not None:
                roles = ["photo" in master.layouts[row.sheet].columns for row in old.rows]
                if len(roles) == 1:
                    issues.append(
                        ImportIssue(
                            "MASTER_MIRROR_REPAIRED",
                            "Неполная пара товара будет восстановлена.",
                            Severity.WARNING,
                            product.article,
                        )
                    )
                elif len(roles) != 2 or roles.count(True) != 1:
                    issues.append(
                        ImportIssue(
                            "MASTER_ARTICLE_CONFLICT",
                            "У товара нет однозначной расчётной/фото пары.",
                            Severity.ERROR,
                            product.article,
                        )
                    )
                if old.photo_error:
                    issues.append(
                        ImportIssue(
                            "MASTER_PHOTO_INVALID",
                            old.photo_error,
                            Severity.ERROR,
                            product.article,
                        )
                    )
        validated.append(replace(item, issues=tuple(issues)))
    selected = SelectLatest().execute(tuple(validated))
    issues = list(selected.issues)
    result = []
    unchanged = 0
    for routed in selected.products:
        p, a = routed.product, routed.assignment
        assert a is not None
        digest = fingerprint_json(
            {
                "values": dict(p.values),
                "supplier": p.supplier,
                "season": a.effective_season,
                "route": a.route.id,
                "sheets": [a.route.calculation_sheet, a.route.photo_sheet],
                "photo": p.photo.pixel_sha256 if p.photo else None,
                "subject": p.attachment.mail.subject,
            }
        )
        old = master.products.get(p.article)
        if old and old.received_at is None and old.received_day is not None:
            if p.attachment.mail.received_at.astimezone().date() < old.received_day:
                unchanged += 1
                continue
        if old and old.received_at is not None:
            if p.attachment.mail.received_at < old.received_at:
                unchanged += 1
                continue
            if p.attachment.mail.received_at == old.received_at:
                if digest == old.fingerprint:
                    unchanged += 1
                else:
                    issues.append(
                        ImportIssue(
                            "MASTER_TIME_CONFLICT",
                            "Дата совпала с сохранённой, данные отличаются.",
                            Severity.ERROR,
                            p.article,
                        )
                    )
                continue
        result.append(ProductWrite(routed, old, digest))
    return WritePlan(
        tuple(result),
        tuple(issues),
        unchanged,
        selected.skipped_articles + len(selected.products) - len(result) - unchanged,
        tuple(validated),
    )
