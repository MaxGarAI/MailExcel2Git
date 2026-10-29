"""Resolve missing destination fields without touching a workbook or the UI."""

from collections.abc import Callable
from dataclasses import dataclass, replace
from pathlib import Path

from kuchenland_importer.application.table_schema import ColumnRegistry, header_key
from kuchenland_importer.application.target_fields import product_fields
from kuchenland_importer.domain.field_review import (
    AddedColumn,
    FieldChoice,
    FieldDecision,
    rule_key,
)
from kuchenland_importer.domain.import_batch import PreparedAttachment
from kuchenland_importer.domain.master import MasterSnapshot
from kuchenland_importer.domain.report import ImportIssue, Severity

Review = Callable[[FieldChoice], FieldDecision]
PROTECTED = frozenset({
    "article", "photo", "season", "mail_received", "mail_subject", "mail_supplier"
})


@dataclass(frozen=True, slots=True)
class FieldReviewResult:
    batch: tuple[PreparedAttachment, ...]
    columns: tuple[AddedColumn, ...]


def review_fields(
    batch: tuple[PreparedAttachment, ...], master: MasterSnapshot, workbook: Path,
    rules: dict[str, FieldDecision], review: Review, registry: ColumnRegistry,
) -> FieldReviewResult:
    groups: dict[tuple[str, bool, str], list[tuple[str, str]]] = {}
    for attachment in batch:
        if not attachment.eligible:
            continue
        for item in attachment.products:
            if item.assignment is None:
                continue
            route = item.assignment.route
            for name, photo in ((route.calculation_sheet, False), (route.photo_sheet, True)):
                if name not in master.layouts:
                    continue
                for key, value in product_fields(item, photo).items():
                    if key not in master.layouts[name].columns and value not in (None, ""):
                        group = (item.product.supplier.strip().casefold(), photo, key)
                        groups.setdefault(group, []).append((name, str(value)[:150]))
    decisions = {}
    added = []
    for (supplier, photo, source), entries in sorted(groups.items()):
        names = tuple(sorted({name for name, _ in entries}))
        common = set.intersection(*(set(master.layouts[n].columns) for n in names))
        targets = tuple(sorted(common - PROTECTED))
        choice = FieldChoice(supplier, photo, source,
                             tuple(dict.fromkeys(v for _, v in entries))[:3], targets, names)
        key = rule_key(workbook, supplier, photo, source)
        decision = rules.get(key)
        if decision is None or decision.action == "map" and decision.target not in targets:
            decision = review(choice)
        if decision.action not in {"skip", "map", "create"}:
            raise ValueError("Неизвестное действие сопоставления.")
        target = decision.target
        if decision.action == "map" and target not in targets:
            raise ValueError("Целевой столбец недоступен или является служебным.")
        if decision.action == "create":
            label = target.strip()
            if (not label or len(label) > 150 or label[0] in "=+-@"
                    or " / " in label or "\n" in label or "\r" in label):
                raise ValueError("Укажите простое название нового столбца до 150 символов.")
            if registry.resolve(label) is not None or header_key(label) in {
                "производитель", "дата", "тема письма"
            }:
                raise ValueError("Такое название уже обозначает стандартное поле; свяжите столбцы.")
            target = "source:" + header_key(label)
            # Never silently map a new title onto a known alias or another column.
            for name in names:
                layout = master.layouts[name]
                if target in layout.columns:
                    continue
                column = max(layout.columns.values()) + 1
                if column > 512:
                    raise ValueError("Достигнут лимит 512 столбцов.")
                layout.columns[target] = column
                added.append(AddedColumn(name, target, label, column))
        rules[key] = decision
        decisions[(supplier, photo, source)] = (decision.action, target)
    prepared = []
    for attachment in batch:
        products = []
        issues = list(attachment.issues)
        for item in attachment.products:
            values = dict(item.product.values)
            if item.assignment is not None:
                supplier = item.product.supplier.strip().casefold()
                for photo in (False, True):
                    fields = product_fields(item, photo)
                    assigned: set[str] = set()
                    for source, value in fields.items():
                        resolved = decisions.get((supplier, photo, source))
                        if resolved is None:
                            continue
                        action, target = resolved
                        if action == "skip" or target == source:
                            continue
                        values[f"review:{int(photo)}:drop:{source}"] = True
                        if target in assigned or target in fields and target != source:
                            issues.append(ImportIssue("FIELD_MAPPING_CONFLICT",
                                f"Несколько полей направлены в {target}.", Severity.ERROR,
                                item.product.article))
                            continue
                        assigned.add(target)
                        values[f"review:{int(photo)}:set:{target}"] = value
            products.append(replace(item, product=replace(item.product, values=values)))
        prepared.append(replace(attachment, products=tuple(products), issues=tuple(issues)))
    return FieldReviewResult(tuple(prepared), tuple(added))
