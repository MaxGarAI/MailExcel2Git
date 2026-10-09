"""Embedded image payloads and failures with source cell coordinates."""

from dataclasses import dataclass
from enum import StrEnum

from kuchenland_importer.domain.product import PhotoAsset


@dataclass(frozen=True, slots=True)
class EmbeddedPhoto:
    sheet: str
    row: int
    column: int
    data: bytes
    source_id: str
    end_row: int | None = None
    crop: tuple[int, int, int, int] = (0, 0, 0, 0)
    rotation: float = 0
    flip_h: bool = False
    flip_v: bool = False
    native: bool = False

    def __post_init__(self) -> None:
        if not (1 <= self.row <= 100000 and 1 <= self.column <= 512):
            raise ValueError("Фото находится за пределами поддерживаемых координат.")
        if self.end_row is not None and not self.row <= self.end_row <= 100000:
            raise ValueError("Недопустимая конечная строка фото.")
        if not self.data or len(self.data) > 20 * 1024 * 1024:
            raise ValueError("Пустое фото или превышен лимит 20 МБ.")


@dataclass(frozen=True, slots=True)
class PhotoProblem:
    sheet: str
    row: int | None
    column: int | None
    message: str


@dataclass(frozen=True, slots=True)
class PhotoExtraction:
    pictures: tuple[EmbeddedPhoto, ...]
    problems: tuple[PhotoProblem, ...] = ()


class PhotoAction(StrEnum):
    ADD = "add"
    REPLACE = "replace"
    KEEP = "keep"
    NONE = "none"


def photo_action(existing: PhotoAsset | None, incoming: PhotoAsset | None) -> PhotoAction:
    if incoming is None:
        return PhotoAction.KEEP if existing is not None else PhotoAction.NONE
    if existing is None:
        return PhotoAction.ADD
    return (
        PhotoAction.KEEP
        if existing.pixel_sha256 == incoming.pixel_sha256
        else PhotoAction.REPLACE
    )
