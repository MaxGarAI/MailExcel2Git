"""Native picture conversion on an owned Excel working copy; never a floating fallback."""

from typing import Any

from PIL import Image

from kuchenland_importer.domain.product import PhotoAsset
from kuchenland_importer.infrastructure.photo_store import pixel_digest


class InCellPictureWriter:
    def place(self, sheet: Any, address: str, asset: PhotoAsset) -> None:
        with Image.open(asset.path) as image:
            if pixel_digest(image) != asset.pixel_sha256:
                raise ValueError("Фото изменилось после подготовки.")
        cell = sheet.Range(address)
        if cell.Count != 1 or cell.MergeCells:
            raise ValueError("Нативное фото требует одной необъединённой ячейки.")
        sheet.Activate()
        cell.Select()
        before = sheet.Shapes.Count
        picture = sheet.Shapes.AddPicture(
            str(asset.path), False, True, cell.Left + 0.1, cell.Top + 0.1,
            max(0.1, cell.Width - 0.2), max(0.1, cell.Height - 0.2),
        )
        try:
            picture.Select()
            picture.PlacePictureInCell()
            if sheet.Shapes.Count != before:
                raise ValueError("Excel не преобразовал фото в изображение в ячейке.")
            # Conversion initializes native images in a fresh book but resamples pixels.
            # Direct insertion then replaces that temporary image with the original PNG.
            cell.Select()
            cell.InsertPictureInCell(str(asset.path))
        except Exception:
            # The caller must discard the working copy if conversion failed.
            if sheet.Shapes.Count > before:
                picture.Delete()
            raise
