from __future__ import annotations

import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from typing import Any

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

NS = {"p": P_NS, "a": A_NS}

TITLE_TYPES = {"title", "ctrTitle"}
BODY_TYPES = {"body"}


@dataclass
class TextStyleStats:
    size_counts: dict[str, int]
    bold_counts: dict[str, int]


def extract_text_style_stats(root: ET.Element) -> TextStyleStats:
    sizes: Counter[str] = Counter()
    bold: Counter[str] = Counter()

    for node in root.findall(".//a:rPr", NS) + root.findall(".//a:defRPr", NS):
        sz = node.attrib.get("sz")
        if sz and sz.isdigit():
            sizes[sz] += 1
        bold_flag = node.attrib.get("b")
        if bold_flag is not None:
            bold[bold_flag] += 1

    return TextStyleStats(size_counts=dict(sizes), bold_counts=dict(bold))


def detect_placeholder_styles(root: ET.Element) -> dict[str, dict[str, Any]]:
    styles: dict[str, dict[str, Any]] = {}
    for shape in root.findall(".//p:sp", NS):
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is None:
            continue
        ph_type = ph.attrib.get("type", "body")
        category = (
            "title"
            if ph_type in TITLE_TYPES
            else "body" if ph_type in BODY_TYPES else None
        )
        if category is None:
            continue
        sizes = []
        bolds = []
        for rpr in shape.findall(".//a:rPr", NS) + shape.findall(".//a:defRPr", NS):
            sz = rpr.attrib.get("sz")
            if sz and sz.isdigit():
                sizes.append(int(sz))
            b = rpr.attrib.get("b")
            if b is not None:
                bolds.append(b)
        if not sizes and not bolds:
            continue
        sizes_counter = Counter(sizes)
        bold_counter = Counter(bolds)
        styles[category] = {
            "size_pt": _sz_to_pt(_most_common(sizes_counter)),
            "bold": _most_common(bold_counter) == "1" if bold_counter else None,
        }
    return styles


def set_layout_text_styles(
    root: ET.Element,
    title_size_pt: float | None,
    title_bold: bool | None,
    body_size_pt: float | None,
    body_bold: bool | None,
) -> int:
    updated = 0
    for shape in root.findall(".//p:sp", NS):
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is None:
            continue
        ph_type = ph.attrib.get("type", "body")
        if ph_type in TITLE_TYPES:
            updated += _apply_shape_style(shape, title_size_pt, title_bold)
        elif ph_type in BODY_TYPES:
            updated += _apply_shape_style(shape, body_size_pt, body_bold)
    return updated


def set_master_text_styles(
    root: ET.Element,
    title_size_pt: float | None,
    title_bold: bool | None,
    body_size_pt: float | None,
    body_bold: bool | None,
) -> int:
    updated = 0
    tx_styles = root.find("p:txStyles", NS)
    if tx_styles is None:
        return 0
    title_style = tx_styles.find("p:titleStyle", NS)
    body_style = tx_styles.find("p:bodyStyle", NS)

    if title_style is not None:
        updated += _apply_level_style(title_style, title_size_pt, title_bold)
    if body_style is not None:
        updated += _apply_level_style(body_style, body_size_pt, body_bold)

    return updated


def _apply_shape_style(
    shape: ET.Element, size_pt: float | None, bold: bool | None
) -> int:
    if size_pt is None and bold is None:
        return 0
    lst_style = shape.find(".//a:lstStyle", NS)
    if lst_style is None:
        tx_body = shape.find(".//a:txBody", NS)
        if tx_body is None:
            return 0
        lst_style = ET.SubElement(tx_body, f"{{{A_NS}}}lstStyle")
    lvl = lst_style.find("a:lvl1pPr", NS)
    if lvl is None:
        lvl = ET.SubElement(lst_style, f"{{{A_NS}}}lvl1pPr")
    def_rpr = lvl.find("a:defRPr", NS)
    if def_rpr is None:
        def_rpr = ET.SubElement(lvl, f"{{{A_NS}}}defRPr")
    return _set_rpr(def_rpr, size_pt, bold)


def _apply_level_style(
    container: ET.Element, size_pt: float | None, bold: bool | None
) -> int:
    lvl = container.find("a:lvl1pPr", NS)
    if lvl is None:
        lvl = ET.SubElement(container, f"{{{A_NS}}}lvl1pPr")
    def_rpr = lvl.find("a:defRPr", NS)
    if def_rpr is None:
        def_rpr = ET.SubElement(lvl, f"{{{A_NS}}}defRPr")
    return _set_rpr(def_rpr, size_pt, bold)


def _set_rpr(node: ET.Element, size_pt: float | None, bold: bool | None) -> int:
    updated = 0
    if size_pt is not None:
        node.set("sz", str(_pt_to_sz(size_pt)))
        updated += 1
    if bold is not None:
        node.set("b", "1" if bold else "0")
        updated += 1
    return updated


def _pt_to_sz(size_pt: float) -> int:
    return int(round(size_pt * 100))


def _sz_to_pt(value: int | None) -> float | None:
    if value is None:
        return None
    return value / 100


def _most_common(counter: Counter) -> Any:
    if not counter:
        return None
    return counter.most_common(1)[0][0]
