from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Iterable

from .package import OOXMLPackage
from .slide_index import slide_parts_in_order

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}


@dataclass
class SanitizeResult:
    slides_updated: int
    clrmap_added: int
    lststyle_added: int
    bg_nofill_added: int


def sanitize_slides(
    pkg: OOXMLPackage, slide_numbers: Iterable[int] | None = None
) -> SanitizeResult:
    slide_parts = slide_parts_in_order(pkg)
    if slide_numbers is not None:
        requested = {num for num in slide_numbers}
        selected = []
        for num in sorted(requested):
            if num < 1 or num > len(slide_parts):
                raise ValueError("Slide number out of range")
            selected.append(slide_parts[num - 1])
        slide_parts = selected

    slides_updated = 0
    clrmap_added = 0
    lststyle_added = 0
    bg_nofill_added = 0

    for slide_part in slide_parts:
        root = ET.fromstring(pkg.read_part(slide_part))
        changed = False

        if _ensure_clrmap_ovr(root):
            clrmap_added += 1
            changed = True
        lst_added = _ensure_lststyle(root)
        if lst_added:
            lststyle_added += lst_added
            changed = True
        if _ensure_bg_nofill(root):
            bg_nofill_added += 1
            changed = True

        if changed:
            pkg.write_part(
                slide_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
            )
            slides_updated += 1

    return SanitizeResult(
        slides_updated=slides_updated,
        clrmap_added=clrmap_added,
        lststyle_added=lststyle_added,
        bg_nofill_added=bg_nofill_added,
    )


def _ensure_clrmap_ovr(root: ET.Element) -> bool:
    if root.find("p:clrMapOvr", NS) is not None:
        return False
    clrmap = ET.Element(f"{{{P_NS}}}clrMapOvr")
    ET.SubElement(clrmap, f"{{{A_NS}}}masterClrMapping")

    transition = root.find("p:transition", NS)
    if transition is not None:
        root.insert(list(root).index(transition), clrmap)
    else:
        root.append(clrmap)
    return True


def _ensure_lststyle(root: ET.Element) -> int:
    added = 0
    for tx_body in root.findall(".//p:txBody", NS):
        if tx_body.find("a:lstStyle", NS) is not None:
            continue
        lst = ET.Element(f"{{{A_NS}}}lstStyle")
        body_pr = tx_body.find("a:bodyPr", NS)
        if body_pr is not None:
            tx_body.insert(list(tx_body).index(body_pr) + 1, lst)
        else:
            tx_body.insert(0, lst)
        added += 1
    return added


def _ensure_bg_nofill(root: ET.Element) -> int:
    bg_pr = root.find("p:cSld/p:bg/p:bgPr", NS)
    if bg_pr is None:
        return 0
    for tag in ["a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:noFill"]:
        if bg_pr.find(tag, NS) is not None:
            return 0
    no_fill = ET.Element(f"{{{A_NS}}}noFill")
    effect = bg_pr.find("a:effectLst", NS)
    if effect is not None:
        bg_pr.insert(list(bg_pr).index(effect), no_fill)
    else:
        bg_pr.append(no_fill)
    return 1
