from __future__ import annotations

from pathlib import Path

from .content_types import ensure_default
from .package import OOXMLPackage

IMAGE_TYPES = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
}


def add_image_part(pkg: OOXMLPackage, image_path: str) -> str:
    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(image_path)
    ext = path.suffix.lower().lstrip(".")
    if ext not in IMAGE_TYPES:
        raise ValueError(f"Unsupported image type: {ext}")

    part_name = _next_media_part(pkg, ext)
    pkg.write_part(part_name, path.read_bytes())
    ensure_default(pkg, ext, IMAGE_TYPES[ext])
    return part_name


def _next_media_part(pkg: OOXMLPackage, ext: str) -> str:
    numbers = []
    for part in pkg.list_parts():
        if not part.startswith("ppt/media/"):
            continue
        name = Path(part).name
        if name.startswith("image") and name.endswith(f".{ext}"):
            raw = name.removeprefix("image").removesuffix(f".{ext}")
            if raw.isdigit():
                numbers.append(int(raw))
    next_index = max(numbers) + 1 if numbers else 1
    return f"ppt/media/image{next_index}.{ext}"
