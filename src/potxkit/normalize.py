from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass

from .package import OOXMLPackage
from .formatting import apply_color_mapping, normalize_mapping
from .slide_index import slide_parts_in_order

SCHEME_SYNONYMS = {
    "dark1": "dk1",
    "light1": "lt1",
    "dark2": "dk2",
    "light2": "lt2",
}


@dataclass
class NormalizeResult:
    slides_total: int
    slides_touched: int
    replacements: int
    per_slide: dict[int, int]


def normalize_slide_colors(
    pkg: OOXMLPackage,
    mapping: dict[str, str],
    slide_numbers: set[int] | None = None,
) -> NormalizeResult:
    slide_parts = slide_parts_in_order(pkg)
    normalized_mapping = normalize_mapping(mapping)

    per_slide: dict[int, int] = {}
    total_replacements = 0
    touched = 0

    for idx, slide_part in enumerate(slide_parts, start=1):
        if slide_numbers and idx not in slide_numbers:
            continue
        root = ET.fromstring(pkg.read_part(slide_part))
        replacements = apply_color_mapping(root, normalized_mapping)
        if replacements:
            pkg.write_part(slide_part, ET.tostring(root, encoding="utf-8", xml_declaration=True))
            total_replacements += replacements
            per_slide[idx] = replacements
            touched += 1

    return NormalizeResult(
        slides_total=len(slide_parts),
        slides_touched=touched,
        replacements=total_replacements,
        per_slide=per_slide,
    )


def parse_slide_numbers(value: str) -> set[int]:
    if not value:
        return set()
    selections: set[int] = set()
    for token in value.split(","):
        token = token.strip()
        if not token:
            continue
        if "-" in token:
            start_str, end_str = token.split("-", 1)
            start = int(start_str)
            end = int(end_str)
            if start <= 0 or end <= 0:
                raise ValueError("Slide numbers must be positive")
            if end < start:
                start, end = end, start
            selections.update(range(start, end + 1))
        else:
            num = int(token)
            if num <= 0:
                raise ValueError("Slide numbers must be positive")
            selections.add(num)
    return selections
