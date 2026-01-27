from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Any, Iterable

from .package import OOXMLPackage
from .rels import parse_relationships, rels_part_for
from .slide_index import slide_parts_in_order

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

NS = {"p": P_NS, "a": A_NS, "r": R_NS}


@dataclass
class DumpTreeOptions:
    include_layout: bool = False
    include_master: bool = False
    include_text: bool = False
    grouped: bool = False


def dump_tree(
    pkg: OOXMLPackage,
    slide_numbers: Iterable[int] | None = None,
    options: DumpTreeOptions | None = None,
) -> dict[str, Any]:
    opts = options or DumpTreeOptions()
    slide_parts = slide_parts_in_order(pkg)

    if slide_numbers is not None:
        requested = {num for num in slide_numbers}
        selected = []
        for num in sorted(requested):
            if num < 1 or num > len(slide_parts):
                raise ValueError("Slide number out of range")
            selected.append(slide_parts[num - 1])
        slide_parts = selected

    slides = []
    for idx, slide_part in enumerate(slide_parts, start=1):
        slide_root = ET.fromstring(pkg.read_part(slide_part))
        layout_part = _slide_layout_part(pkg, slide_part)
        master_part = _layout_master_part(pkg, layout_part) if layout_part else None

        if opts.grouped:
            entry = {
                "slide": idx,
                "part": slide_part,
                "layout": layout_part,
                "master": master_part,
                "local": _collect_layer(slide_root, include_text=opts.include_text),
            }
            if opts.include_layout and layout_part:
                layout_root = ET.fromstring(pkg.read_part(layout_part))
                entry["slideLayout"] = _collect_layer(
                    layout_root, include_text=opts.include_text, part=layout_part
                )
            if opts.include_master and master_part:
                master_root = ET.fromstring(pkg.read_part(master_part))
                entry["slideMaster"] = _collect_layer(
                    master_root, include_text=opts.include_text, part=master_part
                )
            slides.append(entry)
            continue

        slide_entry = {
            "slide": idx,
            "part": slide_part,
            "layout": layout_part,
            "master": master_part,
            "background": _extract_background(slide_root),
            "shapes": _extract_shapes(slide_root, include_text=opts.include_text),
            "has_clrMapOvr": slide_root.find("p:clrMapOvr", NS) is not None,
        }

        if opts.include_layout and layout_part:
            layout_root = ET.fromstring(pkg.read_part(layout_part))
            slide_entry["layout_tree"] = _collect_layer(
                layout_root, include_text=opts.include_text, part=layout_part
            )

        if opts.include_master and master_part:
            master_root = ET.fromstring(pkg.read_part(master_part))
            slide_entry["master_tree"] = _collect_layer(
                master_root, include_text=opts.include_text, part=master_part
            )

        slides.append(slide_entry)

    return {"slides": slides}


def summarize_tree(payload: dict[str, Any], *, local_only: bool = False) -> list[str]:
    lines: list[str] = []
    slides = payload.get("slides", [])
    for slide in slides:
        if local_only and not _slide_has_local_hardcoded(slide):
            continue
        lines.append(f"slide {slide.get('slide')}:")
        for key in ["slideMaster", "slideLayout", "local"]:
            if key not in slide:
                continue
            summary = _summarize_layer(slide[key])
            line = (
                f"  {key}: bg={summary['bg']} "
                f"fills(hard={summary['shape_fill_hard']}, theme={summary['shape_fill_theme']}) "
                f"text(hard={summary['text_color_hard']}, theme={summary['text_color_theme']}) "
                f"fonts={summary['fonts']} sizes={summary['sizes']} "
                f"clrMap={summary['clrmap']}"
            )
            lines.append(line)
    return lines


def _collect_layer(
    root: ET.Element, *, include_text: bool, part: str | None = None
) -> dict[str, Any]:
    data = {
        "part": part,
        "name": _part_basename(part) if part else None,
        "background": _extract_background(root),
        "shapes": _extract_shapes(root, include_text=include_text),
        "has_clrMap": root.find(".//p:clrMap", NS) is not None,
        "has_clrMapOvr": root.find(".//p:clrMapOvr", NS) is not None,
    }
    return data


def _slide_layout_part(pkg: OOXMLPackage, slide_part: str) -> str | None:
    rels_part = rels_part_for(slide_part)
    if not pkg.has_part(rels_part):
        return None
    relationships = parse_relationships(pkg.read_part(rels_part))
    for rel in relationships:
        if rel.type.endswith("/slideLayout"):
            return _resolve_target(posixpath.dirname(slide_part), rel.target)
    return None


def _layout_master_part(pkg: OOXMLPackage, layout_part: str | None) -> str | None:
    if not layout_part:
        return None
    rels_part = rels_part_for(layout_part)
    if not pkg.has_part(rels_part):
        return None
    relationships = parse_relationships(pkg.read_part(rels_part))
    for rel in relationships:
        if rel.type.endswith("/slideMaster"):
            return _resolve_target(posixpath.dirname(layout_part), rel.target)
    return None


def _resolve_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(base_dir, target))


def _extract_background(root: ET.Element) -> dict[str, Any] | None:
    bg_pr = root.find("p:cSld/p:bg/p:bgPr", NS)
    if bg_pr is None:
        return None
    fill = _extract_fill(bg_pr)
    return {"fill": fill} if fill else None


def _extract_shapes(root: ET.Element, *, include_text: bool) -> list[dict[str, Any]]:
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return []
    return [_extract_shape(node, include_text=include_text) for node in list(sp_tree)]


def _extract_shape(node: ET.Element, *, include_text: bool) -> dict[str, Any]:
    tag = _local_name(node.tag)
    if tag == "sp":
        return _extract_sp(node, include_text=include_text)
    if tag == "pic":
        return _extract_pic(node)
    if tag == "graphicFrame":
        return _extract_graphic_frame(node)
    if tag == "grpSp":
        return _extract_group(node, include_text=include_text)
    return {"type": tag}


def _extract_sp(node: ET.Element, *, include_text: bool) -> dict[str, Any]:
    c_nv_pr = node.find("p:nvSpPr/p:cNvPr", NS)
    info = _shape_identity(c_nv_pr)
    ph = node.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is not None:
        info["placeholder"] = {"type": ph.attrib.get("type"), "idx": ph.attrib.get("idx")}
    sp_pr = node.find("p:spPr", NS)
    if sp_pr is not None:
        fill = _extract_fill(sp_pr)
        if fill:
            info["fill"] = fill
    if include_text:
        tx_body = node.find("p:txBody", NS)
        if tx_body is not None:
            info["text"] = _extract_text_info(tx_body)
    return {"type": "shape", **info}


def _extract_pic(node: ET.Element) -> dict[str, Any]:
    c_nv_pr = node.find("p:nvPicPr/p:cNvPr", NS)
    info = _shape_identity(c_nv_pr)
    blip = node.find("p:blipFill/a:blip", NS)
    if blip is not None:
        embed = blip.attrib.get(f"{{{R_NS}}}embed")
        if embed:
            info["embed"] = embed
    fill = _extract_fill(node)
    if fill:
        info["fill"] = fill
    return {"type": "picture", **info}


def _extract_graphic_frame(node: ET.Element) -> dict[str, Any]:
    c_nv_pr = node.find("p:nvGraphicFramePr/p:cNvPr", NS)
    info = _shape_identity(c_nv_pr)
    graphic = node.find("a:graphic/a:graphicData", NS)
    if graphic is not None:
        info["graphic_uri"] = graphic.attrib.get("uri")
    return {"type": "graphicFrame", **info}


def _extract_group(node: ET.Element, *, include_text: bool) -> dict[str, Any]:
    c_nv_pr = node.find("p:nvGrpSpPr/p:cNvPr", NS)
    info = _shape_identity(c_nv_pr)
    children_tree = node.find("p:grpSp/p:spTree", NS)
    children = []
    if children_tree is not None:
        children = [_extract_shape(child, include_text=include_text) for child in list(children_tree)]
    return {"type": "group", **info, "children": children}


def _shape_identity(c_nv_pr: ET.Element | None) -> dict[str, Any]:
    if c_nv_pr is None:
        return {}
    return {"id": c_nv_pr.attrib.get("id"), "name": c_nv_pr.attrib.get("name")}


def _extract_text_info(tx_body: ET.Element) -> dict[str, Any]:
    paragraphs = tx_body.findall("a:p", NS)
    runs = tx_body.findall(".//a:r", NS)
    colors = _extract_color_nodes(tx_body)
    fonts = _extract_text_fonts(tx_body)
    sizes = _extract_text_sizes(tx_body)
    return {
        "paragraphs": len(paragraphs),
        "runs": len(runs),
        "colors": colors,
        "fonts": fonts,
        "sizes_pt": sizes,
        "has_lstStyle": tx_body.find("a:lstStyle", NS) is not None,
    }


def _extract_fill(node: ET.Element) -> dict[str, Any] | None:
    solid = node.find("a:solidFill", NS)
    if solid is not None:
        return {"type": "solid", "color": _extract_color(solid)}

    grad = node.find("a:gradFill", NS)
    if grad is not None:
        stops = []
        for gs in grad.findall("a:gsLst/a:gs", NS):
            stops.append({"pos": gs.attrib.get("pos"), "color": _extract_color(gs)})
        return {"type": "gradient", "stops": stops}

    blip = node.find("a:blipFill", NS)
    if blip is not None:
        return {"type": "image"}

    patt = node.find("a:pattFill", NS)
    if patt is not None:
        return {"type": "pattern", "colors": _extract_color_nodes(patt)}

    if node.find("a:noFill", NS) is not None:
        return {"type": "none"}

    return None


def _extract_color_nodes(node: ET.Element) -> list[dict[str, Any]]:
    colors = []
    for tag in ["srgbClr", "schemeClr", "sysClr", "prstClr"]:
        for color in node.findall(f".//a:{tag}", NS):
            entry = {"kind": tag, "value": color.attrib.get("val")}
            if tag == "sysClr" and "lastClr" in color.attrib:
                entry["lastClr"] = color.attrib.get("lastClr")
            colors.append(entry)
    return colors


def _extract_color(node: ET.Element) -> dict[str, Any] | None:
    for tag in ["srgbClr", "schemeClr", "sysClr", "prstClr"]:
        child = node.find(f"a:{tag}", NS)
        if child is None:
            continue
        entry = {"kind": tag, "value": child.attrib.get("val")}
        if tag == "sysClr" and "lastClr" in child.attrib:
            entry["lastClr"] = child.attrib.get("lastClr")
        return entry
    return None


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _part_basename(part: str | None) -> str | None:
    if not part:
        return None
    return posixpath.basename(part)


def _extract_text_fonts(tx_body: ET.Element) -> list[dict[str, Any]]:
    counts: dict[str, int] = {}
    for rpr in tx_body.findall(".//a:rPr", NS) + tx_body.findall(".//a:defRPr", NS):
        latin = rpr.find("a:latin", NS)
        if latin is None:
            continue
        font = latin.attrib.get("typeface")
        if not font:
            continue
        counts[font] = counts.get(font, 0) + 1
    return [{"value": key, "count": count} for key, count in sorted(counts.items())]


def _extract_text_sizes(tx_body: ET.Element) -> list[dict[str, Any]]:
    counts: dict[float, int] = {}
    for rpr in tx_body.findall(".//a:rPr", NS) + tx_body.findall(".//a:defRPr", NS):
        raw = rpr.attrib.get("sz")
        if not raw or not raw.isdigit():
            continue
        pt = int(raw) / 100
        counts[pt] = counts.get(pt, 0) + 1
    return [
        {"value": size, "count": count}
        for size, count in sorted(counts.items(), key=lambda item: item[0])
    ]


def _summarize_layer(layer: dict[str, Any]) -> dict[str, Any]:
    shapes = layer.get("shapes", [])
    all_shapes = list(_iter_shapes(shapes))
    bg = "none"
    bg_data = layer.get("background") or {}
    bg_fill = bg_data.get("fill")
    if bg_fill:
        bg = _format_fill(bg_fill)

    shape_fill_hard = 0
    shape_fill_theme = 0
    text_color_hard = 0
    text_color_theme = 0
    fonts: dict[str, int] = {}
    sizes: set[float] = set()

    for shape in all_shapes:
        fill = shape.get("fill")
        if fill and fill.get("color"):
            kind = fill["color"].get("kind")
            if kind == "schemeClr":
                shape_fill_theme += 1
            elif kind:
                shape_fill_hard += 1

        text = shape.get("text")
        if not text:
            continue
        for color in text.get("colors", []):
            kind = color.get("kind")
            if kind == "schemeClr":
                text_color_theme += 1
            elif kind:
                text_color_hard += 1
        for font in text.get("fonts", []):
            value = font.get("value")
            count = int(font.get("count", 0))
            if value:
                fonts[value] = fonts.get(value, 0) + count
        for size in text.get("sizes_pt", []):
            value = size.get("value")
            if isinstance(value, (int, float)):
                sizes.add(float(value))

    font_list = [f"{name}({count})" for name, count in sorted(fonts.items())]
    sizes_list = ", ".join(str(int(s) if s.is_integer() else s) for s in sorted(sizes))
    sizes_text = f"{{{sizes_list}}}" if sizes_list else "{}"

    clrmap = "yes" if layer.get("has_clrMap") else "no"
    if layer.get("has_clrMapOvr"):
        clrmap = "override"

    return {
        "bg": bg,
        "shape_fill_hard": shape_fill_hard,
        "shape_fill_theme": shape_fill_theme,
        "text_color_hard": text_color_hard,
        "text_color_theme": text_color_theme,
        "fonts": font_list,
        "sizes": sizes_text,
        "clrmap": clrmap,
    }


def _slide_has_local_hardcoded(slide: dict[str, Any]) -> bool:
    local = slide.get("local") or {}
    bg = local.get("background") or {}
    if _has_hardcoded_fill(bg.get("fill")):
        return True

    for shape in _iter_shapes(local.get("shapes", [])):
        if _has_hardcoded_fill(shape.get("fill")):
            return True
        text = shape.get("text") or {}
        for color in text.get("colors", []):
            kind = color.get("kind")
            if kind and kind != "schemeClr":
                return True
    return False


def _iter_shapes(shapes: Iterable[dict[str, Any]]) -> Iterable[dict[str, Any]]:
    for shape in shapes:
        yield shape
        if shape.get("type") == "group":
            for child in _iter_shapes(shape.get("children", [])):
                yield child


def _format_fill(fill: dict[str, Any]) -> str:
    if not fill:
        return "none"
    if fill.get("type") == "solid":
        color = fill.get("color")
        if not color:
            return "solid"
        kind = color.get("kind")
        val = color.get("value")
        if kind == "schemeClr":
            return f"scheme({val})"
        if kind == "srgbClr":
            return f"srgb(#{val})"
        if kind == "sysClr":
            last = color.get("lastClr")
            suffix = f"/{last}" if last else ""
            return f"sys({val}{suffix})"
        return f"{kind}({val})"
    return fill.get("type", "none")


def _has_hardcoded_fill(fill: dict[str, Any] | None) -> bool:
    if not fill:
        return False
    fill_type = fill.get("type")
    if fill_type == "solid":
        color = fill.get("color") or {}
        kind = color.get("kind")
        return bool(kind and kind != "schemeClr")
    if fill_type == "gradient":
        for stop in fill.get("stops", []):
            color = stop.get("color") or {}
            kind = color.get("kind")
            if kind and kind != "schemeClr":
                return True
        return False
    if fill_type == "pattern":
        for color in fill.get("colors", []):
            kind = color.get("kind")
            if kind and kind != "schemeClr":
                return True
        return False
    return False
