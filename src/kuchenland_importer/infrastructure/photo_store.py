"""Normalize visible pixels, build deterministic collages and persist content-addressed PNGs."""

import io
import math
import struct
import warnings
from hashlib import sha256
from pathlib import Path
from uuid import uuid4

from PIL import Image, ImageOps

from kuchenland_importer.domain.photos import EmbeddedPhoto
from kuchenland_importer.domain.product import PhotoAsset


def pixel_digest(image: Image.Image) -> str:
    rgba = image.convert("RGBA")
    return sha256(
        b"kuchenland-rgba-v1\0" + struct.pack("!II", *rgba.size) + rgba.tobytes()
    ).hexdigest()


def visible_image(photo: EmbeddedPhoto) -> Image.Image:
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        with Image.open(io.BytesIO(photo.data)) as source:
            if source.width * source.height > 40_000_000 or getattr(source, "n_frames", 1) > 1:
                raise ValueError("Фото превышает 40 Мп или содержит несколько кадров.")
            image = ImageOps.exif_transpose(source).convert("RGBA")
    left, top, right, bottom = photo.crop
    if min(photo.crop) < 0 or left + right >= 100000 or top + bottom >= 100000:
        raise ValueError("Обрезка фото не поддерживается или удаляет всё изображение.")
    if any(photo.crop):
        w, h = image.size
        box = (round(w * left / 100000), round(h * top / 100000),
               round(w * (1 - right / 100000)), round(h * (1 - bottom / 100000)))
        if box[0] >= box[2] or box[1] >= box[3]:
            raise ValueError("После обрезки у фото нет пикселей.")
        image = image.crop(box)
    if photo.flip_h:
        image = ImageOps.mirror(image)
    if photo.flip_v:
        image = ImageOps.flip(image)
    if not math.isfinite(photo.rotation):
        raise ValueError("Недопустимый угол поворота фото.")
    if photo.rotation % 360:
        image = image.rotate(-photo.rotation, resample=Image.Resampling.BICUBIC, expand=True)
    return image


class PhotoStore:
    def __init__(self, directory: Path) -> None:
        self.directory = directory.resolve()
        self.directory.mkdir(parents=True, exist_ok=True)

    def build(self, pictures: tuple[EmbeddedPhoto, ...]) -> PhotoAsset:
        if not pictures or len(pictures) > 16:
            raise ValueError("Для одного товара требуется от 1 до 16 изображений.")
        images = {}
        for picture in pictures:
            image = visible_image(picture)
            images[pixel_digest(image)] = image
        ordered = [images[key] for key in sorted(images)]
        if len(ordered) == 1:
            output = ordered[0]
        else:
            columns = math.ceil(math.sqrt(len(ordered)))
            rows = math.ceil(len(ordered) / columns)
            output = Image.new("RGBA", (columns * 512, rows * 512), "white")
            for index, image in enumerate(ordered):
                thumb = ImageOps.contain(image, (496, 496), Image.Resampling.LANCZOS)
                x = index % columns * 512 + (512 - thumb.width) // 2
                y = index // columns * 512 + (512 - thumb.height) // 2
                output.alpha_composite(thumb, (x, y))
        digest = pixel_digest(output)
        path = self.directory / f"{digest}.png"
        if path.exists():
            with Image.open(path) as cached:
                if pixel_digest(cached) != digest:
                    raise ValueError("Контрольная сумма сохранённого фото не совпадает.")
        else:
            partial = self.directory / f"{uuid4().hex}.part"
            try:
                with partial.open("xb") as stream:
                    output.save(stream, format="PNG")
                partial.rename(path)
            finally:
                partial.unlink(missing_ok=True)
        return PhotoAsset(path, digest, output.width, output.height, len(ordered))
