"""Read image parts without opening or saving source books in Excel."""

from pathlib import Path
from uuid import uuid4

from kuchenland_importer.domain.photos import EmbeddedPhoto, PhotoExtraction, PhotoProblem
from kuchenland_importer.infrastructure.excel.drawing_photos import read_drawings
from kuchenland_importer.infrastructure.excel.ooxml_package import OOXMLPackage
from kuchenland_importer.infrastructure.excel.rich_photos import native_images, read_native


class PhotoReader:
    def read(self, path: Path, working_dir: Path | None = None) -> PhotoExtraction:
        if path.suffix.lower() in {".xls", ".xlsb"} and working_dir is not None:
            from kuchenland_importer.infrastructure.excel.legacy_photo_copy import (
                convert_photo_copy,
            )

            working_dir.mkdir(parents=True, exist_ok=True)
            copy = working_dir / f"photo-source-{uuid4().hex}.xlsx"
            try:
                convert_photo_copy(path, copy)
                return self.read(copy)
            finally:
                copy.unlink(missing_ok=True)
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Фото XLS/XLSB требуют конвертации рабочей копии через Excel.")
        package = OOXMLPackage(path)
        try:
            images = native_images(package)
            pictures: list[EmbeddedPhoto] = []
            problems: list[PhotoProblem] = []
            for name, part in package.sheets():
                drawing = read_drawings(package, name, part)
                pictures.extend(drawing.pictures)
                problems.extend(drawing.problems)
                pictures.extend(read_native(package, name, part, images))
            return PhotoExtraction(tuple(pictures), tuple(problems))
        finally:
            package.close()
