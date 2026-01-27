from __future__ import annotations

import copy
import posixpath
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Iterable

from .content_types import ensure_override, remove_override
from .formatting import (
    apply_color_mapping,
    set_text_font_family,
    strip_hardcoded_colors,
    strip_inline_formatting,
)
from .media import add_image_part
from .package import OOXMLPackage
from .rels import (
    Relationship,
    ensure_relationship,
    parse_relationships,
    rels_part_for,
    serialize_relationships,
)
from .slide_index import slide_parts_in_order
from .typography import set_layout_text_styles, set_master_text_styles

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

NS = {"p": P_NS, "a": A_NS, "r": R_NS}

SLIDE_LAYOUT_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
SLIDE_MASTER_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
SLIDE_LAYOUT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
)


def make_layout_from_slide(
    pkg: OOXMLPackage,
    slide_number: int,
    name: str,
    master_index: int = 1,
) -> str:
    slide_parts = slide_parts_in_order(pkg)
    if slide_number < 1 or slide_number > len(slide_parts):
        raise ValueError("Slide number out of range")

    master_part = _master_part_by_index(pkg, master_index)
    template_layout = _first_layout_for_master(pkg, master_part)

    slide_part = slide_parts[slide_number - 1]
    slide_root = ET.fromstring(pkg.read_part(slide_part))
    layout_root = ET.fromstring(pkg.read_part(template_layout))

    slide_c_sld = slide_root.find("p:cSld", NS)
    layout_c_sld = layout_root.find("p:cSld", NS)
    if slide_c_sld is None or layout_c_sld is None:
        raise ValueError("Slide or layout is missing cSld")

    parent_map = {
        child: parent for parent in layout_root.iter() for child in list(parent)
    }
    layout_parent = parent_map.get(layout_c_sld)
    if layout_parent is None:
        raise ValueError("Failed to locate layout cSld parent")

    layout_parent.remove(layout_c_sld)
    layout_parent.insert(0, copy.deepcopy(slide_c_sld))

    layout_root.set("name", name)

    new_layout_part = _next_layout_part(pkg)
    pkg.write_part(
        new_layout_part,
        ET.tostring(layout_root, encoding="utf-8", xml_declaration=True),
    )

    layout_rels = _layout_relationships_from_slide(
        pkg, slide_part, master_part, new_layout_part
    )
    pkg.write_part(rels_part_for(new_layout_part), serialize_relationships(layout_rels))

    master_rel_target = _rel_target(master_part, new_layout_part)
    rel = ensure_relationship(pkg, master_part, SLIDE_LAYOUT_REL, master_rel_target)
    _insert_layout_id(master_part, pkg, rel.id)

    ensure_override(pkg, f"/{new_layout_part}", SLIDE_LAYOUT_CONTENT_TYPE)

    return new_layout_part


def assign_slides_to_layout(
    pkg: OOXMLPackage, slide_numbers: Iterable[int], layout_part: str
) -> None:
    slide_parts = slide_parts_in_order(pkg)
    for number in slide_numbers:
        if number < 1 or number > len(slide_parts):
            raise ValueError("Slide number out of range")
        slide_part = slide_parts[number - 1]
        _set_slide_layout(pkg, slide_part, layout_part)


def apply_palette_to_part(pkg: OOXMLPackage, part: str, mapping: dict[str, str]) -> int:
    root = ET.fromstring(pkg.read_part(part))
    replacements = apply_color_mapping(root, mapping)
    if replacements:
        pkg.write_part(part, ET.tostring(root, encoding="utf-8", xml_declaration=True))
    return replacements


def strip_colors_from_part(pkg: OOXMLPackage, part: str) -> int:
    root = ET.fromstring(pkg.read_part(part))
    removed = strip_hardcoded_colors(root)
    if removed:
        pkg.write_part(part, ET.tostring(root, encoding="utf-8", xml_declaration=True))
    return removed


def strip_fonts_from_part(pkg: OOXMLPackage, part: str) -> int:
    root = ET.fromstring(pkg.read_part(part))
    removed = strip_inline_formatting(root)
    if removed:
        pkg.write_part(part, ET.tostring(root, encoding="utf-8", xml_declaration=True))
    return removed


def set_font_family_for_part(pkg: OOXMLPackage, part: str, font: str) -> int:
    root = ET.fromstring(pkg.read_part(part))
    updated = set_text_font_family(root, font)
    if updated:
        pkg.write_part(part, ET.tostring(root, encoding="utf-8", xml_declaration=True))
    return updated


def set_layout_background_image(
    pkg: OOXMLPackage, layout_part: str, image_path: str
) -> None:
    image_part = add_image_part(pkg, image_path)
    target = _rel_target(layout_part, image_part)
    rel = ensure_relationship(pkg, layout_part, IMAGE_REL, target)

    root = ET.fromstring(pkg.read_part(layout_part))
    c_sld = root.find("p:cSld", NS)
    if c_sld is None:
        raise ValueError("Layout missing cSld")

    bg = c_sld.find("p:bg", NS)
    if bg is None:
        bg = ET.SubElement(c_sld, f"{{{P_NS}}}bg")
    for child in list(bg):
        bg.remove(child)

    bg_pr = ET.SubElement(bg, f"{{{P_NS}}}bgPr")
    blip_fill = ET.SubElement(bg_pr, f"{{{A_NS}}}blipFill")
    blip = ET.SubElement(blip_fill, f"{{{A_NS}}}blip")
    blip.set(f"{{{R_NS}}}embed", rel.id)
    stretch = ET.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    ET.SubElement(stretch, f"{{{A_NS}}}fillRect")

    pkg.write_part(
        layout_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
    )


def add_layout_image_shape(
    pkg: OOXMLPackage,
    layout_part: str,
    image_path: str,
    x: int,
    y: int,
    cx: int,
    cy: int,
    name: str | None = None,
) -> None:
    image_part = add_image_part(pkg, image_path)
    target = _rel_target(layout_part, image_part)
    rel = ensure_relationship(pkg, layout_part, IMAGE_REL, target)

    root = ET.fromstring(pkg.read_part(layout_part))
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        raise ValueError("Layout missing spTree")

    next_id = _next_shape_id(sp_tree)
    pic = ET.SubElement(sp_tree, f"{{{P_NS}}}pic")
    nv_pic_pr = ET.SubElement(pic, f"{{{P_NS}}}nvPicPr")
    c_nv_pr = ET.SubElement(nv_pic_pr, f"{{{P_NS}}}cNvPr")
    c_nv_pr.set("id", str(next_id))
    c_nv_pr.set("name", name or f"Picture {next_id}")
    ET.SubElement(nv_pic_pr, f"{{{P_NS}}}cNvPicPr")
    ET.SubElement(nv_pic_pr, f"{{{P_NS}}}nvPr")

    blip_fill = ET.SubElement(pic, f"{{{P_NS}}}blipFill")
    blip = ET.SubElement(blip_fill, f"{{{A_NS}}}blip")
    blip.set(f"{{{R_NS}}}embed", rel.id)
    stretch = ET.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    ET.SubElement(stretch, f"{{{A_NS}}}fillRect")

    sp_pr = ET.SubElement(pic, f"{{{P_NS}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{A_NS}}}xfrm")
    ET.SubElement(xfrm, f"{{{A_NS}}}off", {"x": str(x), "y": str(y)})
    ET.SubElement(xfrm, f"{{{A_NS}}}ext", {"cx": str(cx), "cy": str(cy)})
    prst = ET.SubElement(sp_pr, f"{{{A_NS}}}prstGeom", {"prst": "rect"})
    ET.SubElement(prst, f"{{{A_NS}}}avLst")

    pkg.write_part(
        layout_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
    )


def set_layout_text_styles_for_part(
    pkg: OOXMLPackage,
    layout_part: str,
    title_size_pt: float | None,
    title_bold: bool | None,
    body_size_pt: float | None,
    body_bold: bool | None,
) -> int:
    root = ET.fromstring(pkg.read_part(layout_part))
    updated = set_layout_text_styles(
        root, title_size_pt, title_bold, body_size_pt, body_bold
    )
    if updated:
        pkg.write_part(
            layout_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
        )
    return updated


def set_master_text_styles_for_part(
    pkg: OOXMLPackage,
    master_part: str,
    title_size_pt: float | None,
    title_bold: bool | None,
    body_size_pt: float | None,
    body_bold: bool | None,
) -> int:
    root = ET.fromstring(pkg.read_part(master_part))
    updated = set_master_text_styles(
        root, title_size_pt, title_bold, body_size_pt, body_bold
    )
    if updated:
        pkg.write_part(
            master_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
        )
    return updated


def resolve_layout_part(pkg: OOXMLPackage, selector: str) -> str:
    if selector.startswith("ppt/slideLayouts/"):
        if not pkg.has_part(selector):
            raise ValueError(f"Layout not found: {selector}")
        return selector

    if selector.isdigit():
        index = int(selector)
        layouts = _layout_parts(pkg)
        if index < 1 or index > len(layouts):
            raise ValueError("Layout index out of range")
        return layouts[index - 1]

    for part in _layout_parts(pkg):
        root = ET.fromstring(pkg.read_part(part))
        name = root.attrib.get("name")
        if name == selector:
            return part

    raise ValueError(f"Layout not found: {selector}")


def resolve_master_part(pkg: OOXMLPackage, selector: str) -> str:
    if selector.startswith("ppt/slideMasters/"):
        if not pkg.has_part(selector):
            raise ValueError(f"Master not found: {selector}")
        return selector
    if selector.isdigit():
        return _master_part_by_index(pkg, int(selector))
    raise ValueError(f"Master not found: {selector}")


def _set_slide_layout(pkg: OOXMLPackage, slide_part: str, layout_part: str) -> None:
    rels_part = rels_part_for(slide_part)
    if not pkg.has_part(rels_part):
        raise ValueError(f"Missing relationships for slide: {slide_part}")
    relationships = parse_relationships(pkg.read_part(rels_part))
    updated = False
    target = _rel_target(slide_part, layout_part)
    for rel in relationships:
        if rel.type == SLIDE_LAYOUT_REL:
            rel.target = target
            updated = True
            break
    if not updated:
        relationships.append(
            Relationship(
                id=_next_rid(relationships), type=SLIDE_LAYOUT_REL, target=target
            )
        )
    pkg.write_part(rels_part, serialize_relationships(relationships))


def _layout_parts(pkg: OOXMLPackage) -> list[str]:
    return sorted(
        [
            p
            for p in pkg.list_parts()
            if p.startswith("ppt/slideLayouts/") and p.endswith(".xml")
        ]
    )


def _master_part_by_index(pkg: OOXMLPackage, index: int) -> str:
    masters = sorted(
        [
            p
            for p in pkg.list_parts()
            if p.startswith("ppt/slideMasters/") and p.endswith(".xml")
        ]
    )
    if not masters:
        raise ValueError("No slide master parts found")
    if index < 1 or index > len(masters):
        raise ValueError("Master index out of range")
    return masters[index - 1]


def _first_layout_for_master(pkg: OOXMLPackage, master_part: str) -> str:
    rels_part = rels_part_for(master_part)
    relationships = parse_relationships(pkg.read_part(rels_part))
    for rel in relationships:
        if rel.type == SLIDE_LAYOUT_REL:
            return _resolve_target(posixpath.dirname(master_part), rel.target)
    layouts = _layout_parts(pkg)
    if not layouts:
        raise ValueError("No slide layout parts found")
    return layouts[0]


def _layout_relationships_from_slide(
    pkg: OOXMLPackage, slide_part: str, master_part: str, layout_part: str
) -> list[Relationship]:
    rels_part = rels_part_for(slide_part)
    slide_rels = parse_relationships(pkg.read_part(rels_part))
    embed_ids = _slide_embed_ids(pkg.read_part(slide_part))

    layout_rels = [rel for rel in slide_rels if rel.id in embed_ids]

    master_target = _rel_target(layout_part, master_part)
    master_id = _next_rid(layout_rels)
    layout_rels.append(
        Relationship(id=master_id, type=SLIDE_MASTER_REL, target=master_target)
    )
    return layout_rels


def _slide_embed_ids(xml_bytes: bytes) -> set[str]:
    root = ET.fromstring(xml_bytes)
    ids = set()
    for node in root.iter():
        for attr in ("embed", "link", "id"):
            rid = node.attrib.get(f"{{{R_NS}}}{attr}")
            if rid:
                ids.add(rid)
    return ids


def _insert_layout_id(master_part: str, pkg: OOXMLPackage, rel_id: str) -> None:
    root = ET.fromstring(pkg.read_part(master_part))
    layout_list = root.find("p:sldLayoutIdLst", NS)
    if layout_list is None:
        layout_list = ET.SubElement(root, f"{{{P_NS}}}sldLayoutIdLst")

    max_id = 0
    for sld_layout_id in layout_list.findall("p:sldLayoutId", NS):
        raw = sld_layout_id.attrib.get("id")
        if raw and raw.isdigit():
            max_id = max(max_id, int(raw))

    new_id = str(max_id + 1 if max_id else 256)
    attrib = {"id": new_id, f"{{{R_NS}}}id": rel_id}
    ET.SubElement(layout_list, f"{{{P_NS}}}sldLayoutId", attrib)

    pkg.write_part(
        master_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
    )


def _next_layout_part(pkg: OOXMLPackage) -> str:
    numbers = []
    for part in _layout_parts(pkg):
        name = posixpath.basename(part)
        if name.startswith("slideLayout") and name.endswith(".xml"):
            raw = name.removeprefix("slideLayout").removesuffix(".xml")
            if raw.isdigit():
                numbers.append(int(raw))
    next_index = max(numbers) + 1 if numbers else 1
    return f"ppt/slideLayouts/slideLayout{next_index}.xml"


def _rel_target(source_part: str, target_part: str) -> str:
    source_dir = posixpath.dirname(source_part)
    return posixpath.relpath(target_part, start=source_dir)


def _resolve_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(base_dir, target))


def _next_rid(relationships: list[Relationship]) -> str:
    existing = {rel.id for rel in relationships}
    idx = 1
    while f"rId{idx}" in existing:
        idx += 1
    return f"rId{idx}"


def slide_size(pkg: OOXMLPackage) -> tuple[int, int]:
    if not pkg.has_part("ppt/presentation.xml"):
        return (0, 0)
    root = ET.fromstring(pkg.read_part("ppt/presentation.xml"))
    sld_sz = root.find("p:sldSz", NS)
    if sld_sz is None:
        return (0, 0)
    cx = int(sld_sz.attrib.get("cx", "0"))
    cy = int(sld_sz.attrib.get("cy", "0"))
    return (cx, cy)


def _next_shape_id(sp_tree: ET.Element) -> int:
    max_id = 0
    for c_nv_pr in sp_tree.findall(".//p:cNvPr", NS):
        raw = c_nv_pr.attrib.get("id")
        if raw and raw.isdigit():
            max_id = max(max_id, int(raw))
    return max_id + 1


@dataclass
class PruneLayoutsResult:
    removed_layouts: list[str]
    unused_layouts: list[str]
    masters_updated: int


@dataclass
class ReindexLayoutsResult:
    layout_mapping: dict[str, str]
    masters_updated: int
    slides_updated: int


def prune_unused_layouts(
    pkg: OOXMLPackage, *, keep_layouts: set[str] | None = None
) -> PruneLayoutsResult:
    keep = keep_layouts or set()
    slide_parts = slide_parts_in_order(pkg)
    used_layouts: set[str] = set()
    for slide_part in slide_parts:
        layout_part = _slide_layout_part(pkg, slide_part)
        if layout_part:
            used_layouts.add(layout_part)

    layout_parts = _layout_parts(pkg)
    unused = [
        layout
        for layout in layout_parts
        if layout not in used_layouts and layout not in keep
    ]

    masters_updated = 0
    for layout in unused:
        masters_updated += _remove_layout_from_masters(pkg, layout)
        pkg.delete_part(layout)
        rels_part = rels_part_for(layout)
        if pkg.has_part(rels_part):
            pkg.delete_part(rels_part)
        remove_override(pkg, layout)

    return PruneLayoutsResult(
        removed_layouts=unused, unused_layouts=unused, masters_updated=masters_updated
    )


def reindex_layouts(pkg: OOXMLPackage) -> ReindexLayoutsResult:
    mapping = _build_layout_reindex_map(pkg)
    if not mapping:
        return ReindexLayoutsResult(
            layout_mapping={}, masters_updated=0, slides_updated=0
        )

    _rename_layout_parts(pkg, mapping)
    slides_updated = _update_slide_layout_relationships(pkg, mapping)
    masters_updated = _reindex_master_layouts(pkg, mapping)

    return ReindexLayoutsResult(
        layout_mapping=mapping,
        masters_updated=masters_updated,
        slides_updated=slides_updated,
    )


def _build_layout_reindex_map(pkg: OOXMLPackage) -> dict[str, str]:
    layout_parts = _layout_parts(pkg)
    if not layout_parts:
        return {}

    mapping: dict[str, str] = {}
    for master_part in _master_parts(pkg):
        order = _master_layout_order(pkg, master_part)
        if not order:
            continue
        for idx, layout_part in enumerate(order, start=1):
            new_name = f"ppt/slideLayouts/slideLayout{idx}.xml"
            if layout_part in mapping and mapping[layout_part] != new_name:
                raise ValueError(
                    f"Layout {layout_part} already mapped to {mapping[layout_part]}"
                )
            mapping[layout_part] = new_name

    return mapping


def _rename_layout_parts(pkg: OOXMLPackage, mapping: dict[str, str]) -> None:
    if not mapping:
        return

    temp_mapping: dict[str, str] = {}
    counter = 1
    for old, new in mapping.items():
        if old == new:
            continue
        temp = f"ppt/slideLayouts/_tmpLayout{counter}.xml"
        while pkg.has_part(temp) or temp in temp_mapping.values():
            counter += 1
            temp = f"ppt/slideLayouts/_tmpLayout{counter}.xml"
        temp_mapping[old] = temp
        counter += 1

    for old, temp in temp_mapping.items():
        pkg.write_part(temp, pkg.read_part(old))
        pkg.delete_part(old)
        old_rels = rels_part_for(old)
        if pkg.has_part(old_rels):
            temp_rels = rels_part_for(temp)
            pkg.write_part(temp_rels, pkg.read_part(old_rels))
            pkg.delete_part(old_rels)
        remove_override(pkg, old)

    for old, new in mapping.items():
        if old == new:
            ensure_override(pkg, new, SLIDE_LAYOUT_CONTENT_TYPE)
            continue
        temp = temp_mapping[old]
        pkg.write_part(new, pkg.read_part(temp))
        pkg.delete_part(temp)
        temp_rels = rels_part_for(temp)
        if pkg.has_part(temp_rels):
            new_rels = rels_part_for(new)
            pkg.write_part(new_rels, pkg.read_part(temp_rels))
            pkg.delete_part(temp_rels)
        ensure_override(pkg, new, SLIDE_LAYOUT_CONTENT_TYPE)


def _update_slide_layout_relationships(
    pkg: OOXMLPackage, mapping: dict[str, str]
) -> int:
    updated = 0
    for slide_part in slide_parts_in_order(pkg):
        rels_part = rels_part_for(slide_part)
        if not pkg.has_part(rels_part):
            continue
        rels = parse_relationships(pkg.read_part(rels_part))
        changed = False
        for rel in rels:
            if not rel.type.endswith("/slideLayout"):
                continue
            target = _resolve_target(posixpath.dirname(slide_part), rel.target)
            if target in mapping:
                rel.target = _rel_target(slide_part, mapping[target])
                changed = True
        if changed:
            pkg.write_part(rels_part, serialize_relationships(rels))
            updated += 1
    return updated


def _reindex_master_layouts(pkg: OOXMLPackage, mapping: dict[str, str]) -> int:
    updated = 0
    for master_part in _master_parts(pkg):
        order = _master_layout_order(pkg, master_part)
        if not order:
            continue

        rels_part = rels_part_for(master_part)
        rels = (
            parse_relationships(pkg.read_part(rels_part))
            if pkg.has_part(rels_part)
            else []
        )
        non_layout = [rel for rel in rels if not rel.type.endswith("/slideLayout")]
        used_ids = {rel.id for rel in non_layout}

        new_layout_rels: list[Relationship] = []
        next_idx = 1
        for layout_part in order:
            new_part = mapping.get(layout_part, layout_part)
            while f"rId{next_idx}" in used_ids:
                next_idx += 1
            rel_id = f"rId{next_idx}"
            used_ids.add(rel_id)
            next_idx += 1
            target = _rel_target(master_part, new_part)
            new_layout_rels.append(
                Relationship(id=rel_id, type=SLIDE_LAYOUT_REL, target=target)
            )

        new_rels = non_layout + new_layout_rels
        pkg.write_part(rels_part, serialize_relationships(new_rels))

        root = ET.fromstring(pkg.read_part(master_part))
        layout_list = root.find("p:sldLayoutIdLst", NS)
        if layout_list is not None:
            for node, layout_rel in zip(
                layout_list.findall("p:sldLayoutId", NS), new_layout_rels
            ):
                node.set(f"{{{R_NS}}}id", layout_rel.id)
        pkg.write_part(
            master_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
        )
        updated += 1
    return updated


def _master_layout_order(pkg: OOXMLPackage, master_part: str) -> list[str]:
    rels_part = rels_part_for(master_part)
    if not pkg.has_part(rels_part):
        return []
    rels = parse_relationships(pkg.read_part(rels_part))
    rel_map = {rel.id: rel for rel in rels}

    root = ET.fromstring(pkg.read_part(master_part))
    layout_list = root.find("p:sldLayoutIdLst", NS)
    if layout_list is None:
        return []

    order: list[str] = []
    for node in layout_list.findall("p:sldLayoutId", NS):
        rid = node.attrib.get(f"{{{R_NS}}}id")
        rel = rel_map.get(rid)
        if rel is None:
            continue
        target = _resolve_target(posixpath.dirname(master_part), rel.target)
        if target:
            order.append(target)
    return order


def _slide_layout_part(pkg: OOXMLPackage, slide_part: str) -> str | None:
    rels_part = rels_part_for(slide_part)
    if not pkg.has_part(rels_part):
        return None
    relationships = parse_relationships(pkg.read_part(rels_part))
    for rel in relationships:
        if rel.type.endswith("/slideLayout"):
            return _resolve_target(posixpath.dirname(slide_part), rel.target)
    return None


def _master_parts(pkg: OOXMLPackage) -> list[str]:
    return sorted(
        [
            p
            for p in pkg.list_parts()
            if p.startswith("ppt/slideMasters/") and p.endswith(".xml")
        ]
    )


def _remove_layout_from_masters(pkg: OOXMLPackage, layout_part: str) -> int:
    updated = 0
    for master_part in _master_parts(pkg):
        rels_part = rels_part_for(master_part)
        if not pkg.has_part(rels_part):
            continue
        relationships = parse_relationships(pkg.read_part(rels_part))
        keep_rels: list[Relationship] = []
        removed_ids: set[str] = set()
        for rel in relationships:
            if rel.type.endswith("/slideLayout"):
                target = _resolve_target(posixpath.dirname(master_part), rel.target)
                if target == layout_part:
                    removed_ids.add(rel.id)
                    continue
            keep_rels.append(rel)

        if not removed_ids:
            continue

        pkg.write_part(rels_part, serialize_relationships(keep_rels))

        root = ET.fromstring(pkg.read_part(master_part))
        layout_list = root.find("p:sldLayoutIdLst", NS)
        if layout_list is not None:
            for layout_id in list(layout_list.findall("p:sldLayoutId", NS)):
                rid = layout_id.attrib.get(f"{{{R_NS}}}id")
                if rid in removed_ids:
                    layout_list.remove(layout_id)
        pkg.write_part(
            master_part, ET.tostring(root, encoding="utf-8", xml_declaration=True)
        )
        updated += 1
    return updated
