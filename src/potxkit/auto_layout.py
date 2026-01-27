from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from .audit import audit_package
from .layout_ops import (
    apply_palette_to_part,
    assign_slides_to_layout,
    make_layout_from_slide,
    strip_colors_from_part,
    strip_fonts_from_part,
)
from .package import OOXMLPackage


@dataclass
class AutoLayoutResult:
    created_layouts: list[str]
    group_count: int


def auto_layout(
    pkg: OOXMLPackage,
    *,
    group_by: Iterable[str] | None = None,
    prefix: str = "Auto Layout",
    master_index: int = 1,
    assign: bool = True,
    strip_colors: bool = False,
    strip_fonts: bool = False,
    palette: dict[str, str] | None = None,
) -> AutoLayoutResult:
    report = audit_package(pkg, group_by=group_by)
    created: list[str] = []

    for idx, group in enumerate(report.groups, start=1):
        slides = group.get("slides", [])
        if not slides:
            continue
        layout_name = f"{prefix} {idx}"
        layout_part = make_layout_from_slide(
            pkg,
            slide_number=slides[0],
            name=layout_name,
            master_index=master_index,
        )
        created.append(layout_part)

        if palette:
            apply_palette_to_part(pkg, layout_part, palette)

        if assign:
            assign_slides_to_layout(pkg, slides, layout_part)

        if strip_colors or strip_fonts:
            for slide_num in slides:
                slide_part = _slide_part_for_number(pkg, slide_num)
                if strip_colors:
                    strip_colors_from_part(pkg, slide_part)
                if strip_fonts:
                    strip_fonts_from_part(pkg, slide_part)

    return AutoLayoutResult(created_layouts=created, group_count=len(report.groups))


def _slide_part_for_number(pkg: OOXMLPackage, slide_number: int) -> str:
    from .slide_index import slide_parts_in_order

    parts = slide_parts_in_order(pkg)
    if slide_number < 1 or slide_number > len(parts):
        raise ValueError("Slide number out of range")
    return parts[slide_number - 1]
