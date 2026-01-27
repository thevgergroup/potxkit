from __future__ import annotations

import xml.etree.ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A_NS, "p": P_NS}

SCHEME_SYNONYMS = {
    "dark1": "dk1",
    "light1": "lt1",
    "dark2": "dk2",
    "light2": "lt2",
}


def apply_color_mapping(root: ET.Element, mapping: dict[str, str]) -> int:
    normalized = normalize_mapping(mapping)
    replacements = 0
    for parent in root.iter():
        children = list(parent)
        for index, child in enumerate(children):
            if child.tag != f"{{{A_NS}}}srgbClr":
                continue
            raw = child.attrib.get("val", "")
            if not raw:
                continue
            key = raw.strip().lstrip("#").upper()
            if key not in normalized:
                continue
            scheme_val = normalized[key]
            scheme = ET.Element(f"{{{A_NS}}}schemeClr", {"val": scheme_val})
            scheme.extend(list(child))
            scheme.text = child.text
            scheme.tail = child.tail
            parent.remove(child)
            parent.insert(index, scheme)
            replacements += 1
    return replacements


def strip_hardcoded_colors(root: ET.Element) -> int:
    removed = 0
    parent_map = _parent_map(root)
    for node in list(root.iter()):
        if node.tag in {f"{{{A_NS}}}srgbClr", f"{{{A_NS}}}sysClr"}:
            parent = parent_map.get(node)
            if parent is not None:
                parent.remove(node)
                removed += 1

    for node in root.findall(".//p:clrMapOvr", NS):
        parent = parent_map.get(node)
        if parent is not None:
            parent.remove(node)
            removed += 1

    for solid in root.findall(".//a:solidFill", NS):
        if not _has_color_child(solid):
            parent = parent_map.get(solid)
            if parent is not None:
                parent.remove(solid)

    for gs in root.findall(".//a:gs", NS):
        if not _has_color_child(gs):
            parent = parent_map.get(gs)
            if parent is not None:
                parent.remove(gs)

    return removed


def strip_inline_formatting(root: ET.Element) -> int:
    removed = 0
    parent_map = _parent_map(root)
    for tag in [
        f"{{{A_NS}}}rPr",
        f"{{{A_NS}}}defRPr",
        f"{{{A_NS}}}lstStyle",
        f"{{{A_NS}}}buClr",
        f"{{{A_NS}}}buSz",
        f"{{{A_NS}}}buFont",
        f"{{{A_NS}}}buChar",
        f"{{{A_NS}}}buAutoNum",
    ]:
        for node in root.findall(f".//{tag}"):
            parent = parent_map.get(node)
            if parent is not None:
                parent.remove(node)
                removed += 1
    return removed


def set_text_font_family(root: ET.Element, typeface: str) -> int:
    updated = 0
    for tag in ["rPr", "defRPr"]:
        for node in root.findall(f".//a:{tag}", NS):
            latin = node.find("a:latin", NS)
            if latin is None:
                latin = ET.SubElement(node, f"{{{A_NS}}}latin")
            latin.set("typeface", typeface)
            updated += 1
    return updated


def normalize_mapping(mapping: dict[str, str]) -> dict[str, str]:
    normalized: dict[str, str] = {}
    for key, value in mapping.items():
        color = key.strip().lstrip("#").upper()
        scheme = value.strip()
        scheme = SCHEME_SYNONYMS.get(scheme.lower(), scheme)
        normalized[color] = scheme
    return normalized


def _parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in list(parent)}


def _has_color_child(node: ET.Element) -> bool:
    for child in list(node):
        if child.tag in {
            f"{{{A_NS}}}srgbClr",
            f"{{{A_NS}}}schemeClr",
            f"{{{A_NS}}}sysClr",
            f"{{{A_NS}}}prstClr",
        }:
            return True
    return False
