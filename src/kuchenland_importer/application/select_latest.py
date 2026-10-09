"""Reject conflicting timestamps and prevent fallback from a rejected newer attachment."""

from collections import defaultdict
from dataclasses import dataclass

from kuchenland_importer.domain.import_batch import PreparedAttachment, SelectedBatch
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.season import RoutedProduct


def same_product(left: RoutedProduct, right: RoutedProduct) -> bool:
    a, b = left.product, right.product
    return (
        a.supplier == b.supplier
        and dict(a.values) == dict(b.values)
        and left.assignment == right.assignment
        and (a.photo.pixel_sha256 if a.photo else None)
        == (b.photo.pixel_sha256 if b.photo else None)
    )


@dataclass(slots=True)
class SelectLatest:
    def execute(self, attachments: tuple[PreparedAttachment, ...]) -> SelectedBatch:
        issues = [issue for item in attachments for issue in item.issues]
        candidates: dict[str, list[PreparedAttachment]] = defaultdict(list)
        for item in attachments:
            for article in item.articles:
                candidates[article].append(item)
        winners = []
        unknown = [
            item.attachment.mail.received_at
            for item in attachments
            if not item.eligible and (not item.articles or not item.identity_complete)
        ]
        for article, files in sorted(candidates.items()):
            newest = max(item.attachment.mail.received_at for item in files)
            final = [item for item in files if item.attachment.mail.received_at == newest]
            if any(not item.eligible for item in final) or any(time >= newest for time in unknown):
                issues.append(
                    ImportIssue(
                        "NEWEST_ATTACHMENT_REJECTED",
                        "Новое вложение ошибочно; старая версия запрещена.",
                        Severity.ERROR,
                        article,
                    )
                )
                continue
            products = [p for item in final for p in item.products if p.product.article == article]
            if not products or any(not same_product(products[0], p) for p in products[1:]):
                issues.append(
                    ImportIssue(
                        "MAIL_TIME_CONFLICT",
                        "Одинаковое время письма, но разные данные товара.",
                        Severity.ERROR,
                        article,
                    )
                )
                continue
            winners.append(products[0])
        return SelectedBatch(tuple(winners), tuple(issues), len(candidates) - len(winners))
