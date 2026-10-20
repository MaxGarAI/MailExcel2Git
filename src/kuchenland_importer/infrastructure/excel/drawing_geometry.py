"""Normalize DrawingML row offsets which may extend beyond the stated anchor row."""

from dataclasses import dataclass
from xml.etree import ElementTree as ET

from kuchenland_importer.infrastructure.excel.ooxml_package import S


@dataclass(frozen=True, slots=True)
class RowGeometry:
    heights: dict[int, int]
    default_height: int

    @classmethod
    def from_sheet(cls, sheet: ET.Element) -> "RowGeometry":
        formatting = sheet.find(f"{{{S}}}sheetFormatPr")
        default = (
            float(formatting.attrib.get("defaultRowHeight", "15")) if formatting is not None else 15
        )
        # Hidden/filter state does not change the SKU to which an anchored image belongs.
        heights = {
            int(r.attrib["r"]): round(float(r.attrib.get("ht", str(default))) * 12700)
            for r in sheet.findall(f"{{{S}}}sheetData/{{{S}}}row")
        }
        return cls(heights, round(default * 12700))

    def normalize(self, row: int, offset: int) -> tuple[int, int]:
        if not 1 <= row <= 100000 or offset < 0:
            raise ValueError("Недопустимая строка/смещение фото.")
        for current in range(row, 100001):
            height = self.heights.get(current, self.default_height)
            if offset < height or height == 0 and offset == 0:
                return current, offset
            offset -= height
        raise ValueError("Фото выходит за пределы поддерживаемых строк.")

    def bounds(self, row: int, offset: int, end_row: int, end_offset: int) -> tuple[int, int]:
        first, top = self.normalize(row, offset)
        last_marker, bottom = self.normalize(end_row, end_offset)
        if (last_marker, bottom) <= (first, top):
            raise ValueError("Конец фото расположен перед началом или высота равна нулю.")
        last = last_marker if bottom > 0 else last_marker - 1
        if last < first:
            raise ValueError("Конец фото расположен перед началом.")
        if last > first:
            first_height = self.heights.get(first, self.default_height) - top
            last_height = bottom if bottom else self.heights.get(last, self.default_height)
            total = (
                first_height
                + last_height
                + sum(self.heights.get(r, self.default_height) for r in range(first + 1, last))
            )
            # Ignore raster/anchor rounding only: at most 4 points AND 2% of image height.
            tolerance = min(4 * 12700, total * 0.02)
            if first_height <= tolerance:
                first += 1
            if last_height <= tolerance:
                last -= 1
        return first, max(first, last)

    def extent_bounds(self, row: int, offset: int, height: int) -> tuple[int, int]:
        if height <= 0:
            raise ValueError("Недопустимая высота фото.")
        first, top = self.normalize(row, offset)
        end, bottom = self.normalize(first, top + height)
        return self.bounds(first, top, end, bottom)
