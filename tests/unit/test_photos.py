import io
from dataclasses import replace
from datetime import UTC, datetime
from pathlib import Path

import pytest
from PIL import Image, UnidentifiedImageError

from kuchenland_importer.application.bind_photos import BindPhotos
from kuchenland_importer.domain.mail import MailMetadata, SavedAttachment
from kuchenland_importer.domain.photos import (
    EmbeddedPhoto,
    PhotoAction,
    PhotoExtraction,
    photo_action,
)
from kuchenland_importer.domain.product import ProductRecord
from kuchenland_importer.domain.workbook import SourceCell, SourceSheet, WorkbookPreview
from kuchenland_importer.infrastructure.excel.columns import load_columns
from kuchenland_importer.infrastructure.photo_store import PhotoStore, pixel_digest, visible_image

CONFIG = Path(__file__).resolve().parents[2] / "config" / "columns.toml"


def payload(color: str = "red", compression: int = 6) -> bytes:
    stream = io.BytesIO()
    Image.new("RGB", (40, 20), color).save(stream, format="PNG", compress_level=compression)
    return stream.getvalue()


def picture(color: str = "red", **kwargs):
    return EmbeddedPhoto("Images", 2, 2, payload(color), "source", **kwargs)


def source_product(path: Path, article: str = "A") -> ProductRecord:
    mail = MailMetadata("entry", "store", "", datetime(2026, 10, 9, tzinfo=UTC), "", "")
    attachment = SavedAttachment(mail, 1, "source.xlsx", path.resolve(), "0" * 64)
    return ProductRecord(article, "Vendor", "Весна", attachment, "Items", 2, {"custom": "keep"})


def photo_sheet() -> SourceSheet:
    return SourceSheet("Images", tuple(tuple(SourceCell(v) for v in r) for r in (
        ("SKU", "Photo", "Name"), ("A", None, "First"), ("B", None, "Second")
    )))


def test_hash_ignores_encoding_and_collage_deduplicates(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    first = store.build((picture(),))
    second = store.build((replace(picture(), data=payload(compression=0)),))
    duplicate = store.build((picture(), picture()))
    assert first == second == duplicate
    assert not list(tmp_path.glob("*.part"))


def test_collage_is_independent_of_drawing_order(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    result = store.build((picture("red"), picture("blue")))
    reverse = store.build((picture("blue"), picture("red")))
    assert result == reverse and (result.width, result.height) == (1024, 512)
    assert result.image_count == 2
    with Image.open(result.path) as image:
        assert len(image.getcolors(maxcolors=10000)) >= 3  # Both images and white margins.


def test_visible_crop_and_rotation_change_pixels() -> None:
    original = picture()
    cropped = visible_image(replace(original, crop=(50000, 0, 0, 0)))
    rotated = visible_image(replace(original, rotation=90))
    assert cropped.size == (20, 20) and rotated.size == (20, 40)
    assert pixel_digest(cropped) != pixel_digest(visible_image(original))


@pytest.mark.parametrize("crop", [(100000, 0, 0, 0), (-1, 0, 0, 0), (50000, 0, 50000, 0)])
def test_invalid_crop_is_not_silently_ignored(crop: tuple[int, int, int, int]) -> None:
    with pytest.raises(ValueError, match="Обрезка"):
        visible_image(picture(crop=crop))


def test_corrupt_payload_and_cache_are_detected(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    with pytest.raises(UnidentifiedImageError):
        store.build((replace(picture(), data=b"broken"),))
    asset = store.build((picture(),))
    Image.new("RGB", (40, 20), "blue").save(asset.path)
    with pytest.raises(ValueError, match="сумма"):
        store.build((picture(),))


@pytest.mark.parametrize("count", [0, 17])
def test_photo_count_limit(tmp_path: Path, count: int) -> None:
    with pytest.raises(ValueError, match="от 1 до 16"):
        PhotoStore(tmp_path).build((picture(),) * count)


def test_photo_update_policy(tmp_path: Path) -> None:
    store = PhotoStore(tmp_path)
    old = store.build((picture("red"),))
    new = store.build((picture("blue"),))
    assert photo_action(None, None) == PhotoAction.NONE
    assert photo_action(None, new) == PhotoAction.ADD
    assert photo_action(old, old) == PhotoAction.KEEP
    assert photo_action(old, None) == PhotoAction.KEEP
    assert photo_action(old, new) == PhotoAction.REPLACE


def test_binding_uses_article_and_photo_column(tmp_path: Path) -> None:
    first = source_product(tmp_path / "book.xlsx")
    second = source_product(tmp_path / "book.xlsx", "B")
    # Source products and picture records have unrelated orders.
    pictures = (replace(picture("blue"), row=3), picture(), replace(picture("green"), column=3))
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((second, first), ()), (photo_sheet(),), PhotoExtraction(pictures)
    )
    assert not result.issues
    assert all(p.photo is not None for p in result.products)
    assert result.products[0].photo.pixel_sha256 != result.products[1].photo.pixel_sha256
    assert dict(result.products[0].values) == {"custom": "keep"}


def test_photo_spanning_products_blocks_attachment(tmp_path: Path) -> None:
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((source_product(tmp_path / "book.xlsx"),), ()),
        (photo_sheet(),), PhotoExtraction((picture(end_row=3),)),
    )
    assert result.issues[0].code == "PHOTO_SPANS_PRODUCTS"
    assert result.products[0].photo is None


def test_missing_photo_is_a_valid_absence(tmp_path: Path) -> None:
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((source_product(tmp_path / "book.xlsx"),), ()),
        (photo_sheet(),), PhotoExtraction(()),
    )
    assert not result.issues and result.products[0].photo is None


def test_corrupt_photo_becomes_explicit_error(tmp_path: Path) -> None:
    result = BindPhotos(load_columns(CONFIG), PhotoStore(tmp_path / "photos")).execute(
        WorkbookPreview((source_product(tmp_path / "book.xlsx"),), ()),
        (photo_sheet(),), PhotoExtraction((replace(picture(), data=b"broken"),)),
    )
    assert result.issues[0].code == "PHOTO_INVALID"
    assert result.products[0].photo is None
