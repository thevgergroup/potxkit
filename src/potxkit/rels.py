from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
from dataclasses import dataclass

from .package import OOXMLPackage

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass
class Relationship:
    id: str
    type: str
    target: str
    target_mode: str | None = None


def rels_part_for(source_part: str) -> str:
    source = _strip_leading_slash(source_part)
    if source == "":
        return "_rels/.rels"
    source_dir = posixpath.dirname(source)
    source_base = posixpath.basename(source)
    return posixpath.join(source_dir, "_rels", f"{source_base}.rels")


def source_part_for(rels_part: str) -> str:
    rels = _strip_leading_slash(rels_part)
    if rels == "_rels/.rels":
        return ""
    rels_dir = posixpath.dirname(rels)
    source_dir = posixpath.dirname(rels_dir)
    source_base = posixpath.basename(rels).removesuffix(".rels")
    return posixpath.join(source_dir, source_base)


def parse_relationships(xml_bytes: bytes) -> list[Relationship]:
    root = ET.fromstring(xml_bytes)
    relationships = []
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        relationships.append(
            Relationship(
                id=rel.attrib.get("Id", ""),
                type=rel.attrib.get("Type", ""),
                target=rel.attrib.get("Target", ""),
                target_mode=rel.attrib.get("TargetMode"),
            )
        )
    return relationships


def serialize_relationships(relationships: list[Relationship]) -> bytes:
    root = ET.Element(f"{{{REL_NS}}}Relationships")
    for rel in relationships:
        attrib = {"Id": rel.id, "Type": rel.type, "Target": rel.target}
        if rel.target_mode:
            attrib["TargetMode"] = rel.target_mode
        ET.SubElement(root, f"{{{REL_NS}}}Relationship", attrib)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def get_relationships(pkg: OOXMLPackage, source_part: str) -> list[Relationship]:
    rels_part = rels_part_for(source_part)
    if not pkg.has_part(rels_part):
        return []
    return parse_relationships(pkg.read_part(rels_part))


def write_relationships(
    pkg: OOXMLPackage, source_part: str, relationships: list[Relationship]
) -> None:
    rels_part = rels_part_for(source_part)
    pkg.write_part(rels_part, serialize_relationships(relationships))


def ensure_relationship(
    pkg: OOXMLPackage, source_part: str, rel_type: str, target: str
) -> Relationship:
    relationships = get_relationships(pkg, source_part)
    for rel in relationships:
        if rel.type == rel_type and rel.target == target:
            return rel

    next_id = _next_rid(relationships)
    new_rel = Relationship(id=next_id, type=rel_type, target=target)
    relationships.append(new_rel)
    write_relationships(pkg, source_part, relationships)
    return new_rel


def _next_rid(relationships: list[Relationship]) -> str:
    existing = {rel.id for rel in relationships}
    index = 1
    while f"rId{index}" in existing:
        index += 1
    return f"rId{index}"


def _strip_leading_slash(path: str) -> str:
    return path[1:] if path.startswith("/") else path
