"""A reviewable list of validated products; saved counts are produced after commit."""

from dataclasses import dataclass

from kuchenland_importer.domain.import_batch import PreparedAttachment
from kuchenland_importer.domain.master import ExistingProduct
from kuchenland_importer.domain.report import ImportIssue
from kuchenland_importer.domain.season import RoutedProduct


@dataclass(frozen=True, slots=True)
class ProductWrite:
    incoming: RoutedProduct
    previous: ExistingProduct | None
    fingerprint: str


@dataclass(frozen=True, slots=True)
class WritePlan:
    products: tuple[ProductWrite, ...]
    issues: tuple[ImportIssue, ...]
    unchanged: int
    skipped: int
    attachments: tuple[PreparedAttachment, ...] = ()
