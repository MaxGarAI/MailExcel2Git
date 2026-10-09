"""Prepared attachments retain rejected article identities for newest-mail selection."""

from dataclasses import dataclass

from kuchenland_importer.domain.mail import SavedAttachment
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.season import RoutedProduct


@dataclass(frozen=True, slots=True)
class PreparedAttachment:
    attachment: SavedAttachment
    products: tuple[RoutedProduct, ...]
    articles: frozenset[str]
    issues: tuple[ImportIssue, ...]
    identity_complete: bool = True

    @property
    def eligible(self) -> bool:
        return not any(issue.severity == Severity.ERROR for issue in self.issues)


@dataclass(frozen=True, slots=True)
class SelectedBatch:
    products: tuple[RoutedProduct, ...]
    issues: tuple[ImportIssue, ...]
    skipped_articles: int
