"""Read saved Top/Bottom filter parameters which the Excel COM getter does not expose."""

from pathlib import Path

from kuchenland_importer.infrastructure.excel.ooxml_package import OOXMLPackage, S
from kuchenland_importer.infrastructure.excel.read_limits import MASTER_LIMITS


def read_top_filters(path: Path) -> dict[tuple[str, int], tuple[int, str]]:
    result = {}
    package = OOXMLPackage(path, limits=MASTER_LIMITS)
    try:
        for name, part in package.sheets():
            for column in package.xml(part).findall(f"{{{S}}}autoFilter/{{{S}}}filterColumn"):
                top = column.find(f"{{{S}}}top10")
                if top is None:
                    continue
                is_top = top.attrib.get("top", "1") in {"1", "true"}
                percent = top.attrib.get("percent", "0") in {"1", "true"}
                operator = (3 if is_top else 5) + int(percent)
                result[(name, int(column.attrib["colId"]) + 1)] = (operator, top.attrib["val"])
    finally:
        package.close()
    return result
