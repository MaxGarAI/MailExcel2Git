from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image

from kuchenland_importer.domain.photos import EmbeddedPhoto
from kuchenland_importer.infrastructure.excel.incell_writer import InCellPictureWriter
from kuchenland_importer.infrastructure.photo_store import PhotoStore


def asset_at(root: Path):
    file = root / "image.png"
    Image.new("RGB", (40, 30), "red").save(file)
    return PhotoStore(root / "photos").build((
        EmbeddedPhoto("Images", 2, 2, file.read_bytes(), "QA"),
    ))


def test_changed_photo_is_rejected_before_excel_access(tmp_path: Path) -> None:
    asset = asset_at(tmp_path)
    Image.new("RGB", (40, 30), "blue").save(asset.path)
    with pytest.raises(ValueError, match="изменилось"):
        InCellPictureWriter().place(None, "B2", asset)


def test_merged_destination_is_rejected(tmp_path: Path) -> None:
    sheet = SimpleNamespace(Range=lambda _: SimpleNamespace(Count=1, MergeCells=True))
    with pytest.raises(ValueError, match="необъединённой"):
        InCellPictureWriter().place(sheet, "B2", asset_at(tmp_path))


def test_noop_conversion_raises_and_removes_temporary_floating_picture(tmp_path: Path) -> None:
    asset = asset_at(tmp_path)
    shapes = SimpleNamespace(Count=0)
    deleted = []

    def add(*args):
        shapes.Count += 1
        return SimpleNamespace(
            Select=lambda: None, PlacePictureInCell=lambda: None,
            Delete=lambda: deleted.append(True),
        )

    shapes.AddPicture = add
    cell = SimpleNamespace(Count=1, MergeCells=False, Left=0, Top=0, Width=30, Height=20,
                           Select=lambda: None)
    sheet = SimpleNamespace(Range=lambda _: cell, Shapes=shapes, Activate=lambda: None)
    with pytest.raises(ValueError, match="не преобразовал"):
        InCellPictureWriter().place(sheet, "B2", asset)
    assert deleted == [True]
