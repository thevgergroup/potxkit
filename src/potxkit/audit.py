from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from typing import Any, Iterable

from .package import OOXMLPackage
from .rels import parse_relationships, rels_part_for
from .slide_index import slide_parts_in_order
from .typography import extract_text_style_stats

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

NS = {"p": P_NS, "a": A_NS}


@dataclass
class AuditReport:
    slides_total: int
    slides_audited: int
    per_slide: dict[int, dict[str, Any]]
    masters: dict[str, dict[str, Any]]
    layouts: dict[str, dict[str, Any]]
    groups: list[dict[str, Any]]
    theme: dict[str, str] | None
    group_by: list[str]


def audit_package(
    pkg: OOXMLPackage,
    slide_numbers: set[int] | None = None,
    group_by: Iterable[str] | None = None,
) -> AuditReport:
    slide_parts = slide_parts_in_order(pkg)
    per_slide: dict[int, dict[str, Any]] = {}
    masters = _summarize_parts(pkg, "ppt/slideMasters/")
    layouts = _summarize_parts(pkg, "ppt/slideLayouts/")
    group_by_list = _normalize_group_by(group_by)

    layout_master = _layout_master_map(pkg)

    for index, slide_part in enumerate(slide_parts, start=1):
        if slide_numbers and index not in slide_numbers:
            continue
        root = ET.fromstring(pkg.read_part(slide_part))
        slide_layout = _slide_layout_part(pkg, slide_part)
        master_part = layout_master.get(slide_layout) if slide_layout else None

        per_slide[index] = {
            "slide_part": slide_part,
            "layout_part": slide_layout,
            "master_part": master_part,
            "color_counts": _color_counts(root),
            "shape_colors": _shape_color_counts(root),
            "text_colors": _text_color_counts(root),
            "text_styles": _text_style_summary(root),
            "has_clrMapOvr": root.find(".//p:clrMapOvr", NS) is not None,
            "background": _background_flags(root),
            "fills": _fill_counts(root),
            "pictures": len(root.findall(".//p:pic", NS)),
            "top_srgb": _top_srgb(root),
        }

    return AuditReport(
        slides_total=len(slide_parts),
        slides_audited=len(per_slide),
        per_slide=per_slide,
        masters=masters,
        layouts=layouts,
        groups=_group_slides(per_slide, group_by_list),
        theme=_theme_summary(pkg),
        group_by=group_by_list,
    )


def _color_counts(root: ET.Element) -> dict[str, int]:
    return {
        "srgb": len(root.findall(".//a:srgbClr", NS)),
        "scheme": len(root.findall(".//a:schemeClr", NS)),
        "sysclr": len(root.findall(".//a:sysClr", NS)),
    }


def _fill_counts(root: ET.Element) -> dict[str, int]:
    return {
        "solid": len(root.findall(".//a:solidFill", NS)),
        "grad": len(root.findall(".//a:gradFill", NS)),
        "blip": len(root.findall(".//a:blipFill", NS)),
    }


def _shape_color_counts(root: ET.Element) -> dict[str, int]:
    return {
        "srgb": len(root.findall(".//p:spPr//a:srgbClr", NS)),
        "scheme": len(root.findall(".//p:spPr//a:schemeClr", NS)),
        "sysclr": len(root.findall(".//p:spPr//a:sysClr", NS)),
    }


def _text_color_counts(root: ET.Element) -> dict[str, int]:
    return {
        "srgb": len(
            root.findall(".//a:rPr//a:srgbClr", NS)
            + root.findall(".//a:defRPr//a:srgbClr", NS)
            + root.findall(".//a:lstStyle//a:srgbClr", NS)
            + root.findall(".//a:buClr//a:srgbClr", NS)
        ),
        "scheme": len(
            root.findall(".//a:rPr//a:schemeClr", NS)
            + root.findall(".//a:defRPr//a:schemeClr", NS)
            + root.findall(".//a:lstStyle//a:schemeClr", NS)
            + root.findall(".//a:buClr//a:schemeClr", NS)
        ),
        "sysclr": len(
            root.findall(".//a:rPr//a:sysClr", NS)
            + root.findall(".//a:defRPr//a:sysClr", NS)
            + root.findall(".//a:lstStyle//a:sysClr", NS)
            + root.findall(".//a:buClr//a:sysClr", NS)
        ),
    }


def _text_style_summary(root: ET.Element) -> dict[str, Any]:
    stats = extract_text_style_stats(root)
    size_counter = Counter(
        {int(k): v for k, v in stats.size_counts.items() if str(k).isdigit()}
    )
    bold_counter = Counter(stats.bold_counts)
    sizes = [
        {"pt": size / 100, "count": count}
        for size, count in size_counter.most_common(5)
    ]
    return {
        "top_sizes": sizes,
        "bold": dict(bold_counter),
    }


def _background_flags(root: ET.Element) -> dict[str, bool]:
    bg = root.find("p:cSld/p:bg", NS)
    bg_pr = bg.find("p:bgPr", NS) if bg is not None else None
    bg_ref = bg.find("p:bgRef", NS) if bg is not None else None

    return {
        "bg_ref": bg_ref is not None,
        "bg_blip": bg_pr is not None and bg_pr.find("a:blipFill", NS) is not None,
        "bg_grad": bg_pr is not None and bg_pr.find("a:gradFill", NS) is not None,
        "bg_solid": bg_pr is not None and bg_pr.find("a:solidFill", NS) is not None,
    }


def _top_srgb(root: ET.Element, limit: int = 5) -> list[dict[str, int]]:
    values = [
        node.attrib.get("val", "").upper()
        for node in root.findall(".//a:srgbClr", NS)
    ]
    values = [val for val in values if val]
    counter = Counter(values)
    return [
        {"value": color, "count": count}
        for color, count in counter.most_common(limit)
    ]


def _slide_layout_part(pkg: OOXMLPackage, slide_part: str) -> str | None:
    rels_part = rels_part_for(slide_part)
    if not pkg.has_part(rels_part):
        return None
    relationships = parse_relationships(pkg.read_part(rels_part))
    for rel in relationships:
        if rel.type.endswith("/slideLayout"):
            return _resolve_target(posixpath.dirname(slide_part), rel.target)
    return None


def _layout_master_map(pkg: OOXMLPackage) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for part in pkg.list_parts():
        if not part.startswith("ppt/slideLayouts/") or not part.endswith(".xml"):
            continue
        rels_part = rels_part_for(part)
        if not pkg.has_part(rels_part):
            continue
        relationships = parse_relationships(pkg.read_part(rels_part))
        for rel in relationships:
            if rel.type.endswith("/slideMaster"):
                mapping[part] = _resolve_target(posixpath.dirname(part), rel.target)
                break
    return mapping


def _summarize_parts(pkg: OOXMLPackage, prefix: str) -> dict[str, dict[str, Any]]:
    summary: dict[str, dict[str, Any]] = {}
    for part in pkg.list_parts():
        if not part.startswith(prefix) or not part.endswith(".xml"):
            continue
        root = ET.fromstring(pkg.read_part(part))
        summary[part] = {
            "color_counts": _color_counts(root),
            "shape_colors": _shape_color_counts(root),
            "text_colors": _text_color_counts(root),
            "fills": _fill_counts(root),
            "pictures": len(root.findall(".//p:pic", NS)),
            "top_srgb": _top_srgb(root),
            "has_clrMap": root.find(".//p:clrMap", NS) is not None,
            "has_clrMapOvr": root.find(".//p:clrMapOvr", NS) is not None,
        }
    return summary


def _group_slides(
    per_slide: dict[int, dict[str, Any]], group_by: list[str]
) -> list[dict[str, Any]]:
    groups: dict[tuple[object, ...], dict[str, Any]] = {}
    for slide_num, data in per_slide.items():
        palette = tuple(entry["value"] for entry in data.get("top_srgb", []))
        key_parts: list[object] = []
        if "l" in group_by:
            key_parts.append(data.get("layout_part"))
            key_parts.append(data.get("master_part"))
        if "b" in group_by:
            key_parts.append(_background_signature(data))
        if "p" in group_by:
            key_parts.append(palette)
        key = tuple(key_parts)
        group = groups.setdefault(
            key,
            {
                "layout_part": data.get("layout_part"),
                "master_part": data.get("master_part"),
                "background": _background_signature(data),
                "palette": list(palette),
                "slides": [],
                "hardcoded_total": 0,
                "text_srgb_total": 0,
                "shape_srgb_total": 0,
                "clrMapOvr_slides": 0,
                "image_slides": 0,
                "custom_bg_slides": 0,
            },
        )
        group["slides"].append(slide_num)
        counts = data["color_counts"]
        text_counts = data["text_colors"]
        shape_counts = data["shape_colors"]
        fills = data["fills"]
        bg = data["background"]

        group["hardcoded_total"] += counts["srgb"] + counts["sysclr"]
        group["text_srgb_total"] += text_counts["srgb"]
        group["shape_srgb_total"] += shape_counts["srgb"]
        if data.get("has_clrMapOvr"):
            group["clrMapOvr_slides"] += 1
        if data.get("pictures") or fills["blip"]:
            group["image_slides"] += 1
        if bg["bg_blip"] or bg["bg_grad"] or bg["bg_solid"] or bg["bg_ref"]:
            group["custom_bg_slides"] += 1

    for group in groups.values():
        group["slides"].sort()
    return list(groups.values())


def _background_signature(data: dict[str, Any]) -> str:
    bg = data.get("background", {})
    flags = []
    if bg.get("bg_blip"):
        flags.append("blip")
    if bg.get("bg_grad"):
        flags.append("grad")
    if bg.get("bg_solid"):
        flags.append("solid")
    if bg.get("bg_ref"):
        flags.append("ref")
    return "+".join(flags) if flags else "none"


def _normalize_group_by(value: Iterable[str] | None) -> list[str]:
    if value is None:
        return ["p", "l"]
    selected = []
    for token in value:
        if token not in {"p", "b", "l"}:
            raise ValueError(f"Invalid group-by option: {token}")
        if token not in selected:
            selected.append(token)
    return selected


def _theme_summary(pkg: OOXMLPackage) -> dict[str, str] | None:
    theme_parts = [
        part
        for part in pkg.list_parts()
        if part.startswith("ppt/theme/") and part.endswith(".xml")
    ]
    if not theme_parts:
        return None
    non_override = [part for part in theme_parts if "themeOverride" not in part]
    theme_part = sorted(non_override or theme_parts)[0]
    root = ET.fromstring(pkg.read_part(theme_part))
    theme_elements = root.find("a:themeElements", NS)
    clr_scheme = theme_elements.find("a:clrScheme", NS) if theme_elements is not None else None
    font_scheme = theme_elements.find("a:fontScheme", NS) if theme_elements is not None else None
    return {
        "part": theme_part,
        "theme_name": root.attrib.get("name", ""),
        "color_scheme_name": clr_scheme.attrib.get("name", "") if clr_scheme is not None else "",
        "font_scheme_name": font_scheme.attrib.get("name", "") if font_scheme is not None else "",
    }


def _resolve_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(base_dir, target))
