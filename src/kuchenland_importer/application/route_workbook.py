"""Keep row diagnostics while making attachment eligibility an all-or-nothing decision."""

from dataclasses import dataclass

from kuchenland_importer.application.resolve_season import ResolveSeason
from kuchenland_importer.domain.errors import SeasonRoutingError
from kuchenland_importer.domain.report import ImportIssue, Severity
from kuchenland_importer.domain.season import RoutedProduct, RoutedWorkbook
from kuchenland_importer.domain.workbook import WorkbookPreview


@dataclass(slots=True)
class RouteWorkbook:
    resolver: ResolveSeason

    def execute(self, preview: WorkbookPreview) -> RoutedWorkbook:
        issues = list(preview.issues)
        products = []
        for product in preview.products:
            try:
                assignment = self.resolver.execute(product)
                issues.extend(assignment.issues)
            except SeasonRoutingError as error:
                assignment = None
                issues.append(
                    ImportIssue(
                        error.code,
                        str(error),
                        Severity.ERROR,
                        f"{product.source_sheet}:{product.source_row}",
                    )
                )
            products.append(RoutedProduct(product, assignment))
        return RoutedWorkbook(tuple(products), tuple(issues))
