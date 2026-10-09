"""Prepare typed products directly from verified originals, without lossy JSON round-trips."""

from pathlib import Path

from kuchenland_importer.application.bind_photos import BindPhotos
from kuchenland_importer.application.normalize_workbook import NormalizeWorkbook, article_value
from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.application.route_workbook import RouteWorkbook
from kuchenland_importer.application.table_schema import ColumnRegistry, find_header
from kuchenland_importer.domain.import_batch import PreparedAttachment
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.workbook import SourceSheet
from kuchenland_importer.infrastructure.capture_loader import load_attachments
from kuchenland_importer.infrastructure.excel.photo_reader import PhotoReader
from kuchenland_importer.infrastructure.excel.reader import WorkbookReader
from kuchenland_importer.infrastructure.photo_store import PhotoStore


def known_articles(
    sheets: tuple[SourceSheet, ...],
    registry: ColumnRegistry,
) -> tuple[frozenset[str], bool]:
    articles: set[str] = set()
    complete = True
    for sheet in sheets:
        try:
            header = find_header(sheet, registry)
        except ValueError:
            complete = False
            continue
        if header is None or not {"article", "supplier", "season"} <= set(header.columns):
            continue
        column = header.columns.index("article")
        for row in sheet.rows[header.row + header.depth - 1 :]:
            try:
                articles.add(article_value(row[column]))
            except ValueError:
                if row[column].value not in {None, ""}:
                    complete = False
                continue
    return frozenset(articles), complete


def prepare_batch(
    manifest: Path,
    registry: ColumnRegistry,
    resolver: ResolveSeason,
    directory: Path,
) -> tuple[PreparedAttachment, ...]:
    result = []
    store = PhotoStore(directory)
    for attachment in load_attachments(manifest):
        identities: frozenset[str] = frozenset()
        complete = False
        try:
            sheets = WorkbookReader().read(attachment.path)
            identities, complete = known_articles(sheets, registry)
            preview = NormalizeWorkbook(registry).execute(attachment, sheets)
            preview = BindPhotos(registry, store).execute(
                preview,
                sheets,
                PhotoReader().read(attachment.path, directory),
            )
            routed = RouteWorkbook(resolver).execute(preview)
            result.append(
                PreparedAttachment(attachment, routed.products, identities, routed.issues, complete)
            )
        except Exception as error:
            result.append(
                PreparedAttachment(
                    attachment,
                    (),
                    identities,
                    (
                        ImportIssue(
                            "WORKBOOK_READ_FAILED",
                            str(error),
                            Severity.ERROR,
                            attachment.original_name,
                        ),
                    ),
                    complete,
                )
            )
    return tuple(result)
