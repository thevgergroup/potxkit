from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET

from .package import OOXMLPackage

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"p": P_NS, "r": R_NS}


def slide_parts_in_order(pkg: OOXMLPackage) -> list[str]:
    if not pkg.has_part("ppt/presentation.xml") or not pkg.has_part(
        "ppt/_rels/presentation.xml.rels"
    ):
        return _fallback_slide_parts(pkg)

    presentation = ET.fromstring(pkg.read_part("ppt/presentation.xml"))
    rels = _read_rels(pkg, "ppt/_rels/presentation.xml.rels")

    slide_parts = []
    for sld_id in presentation.findall("p:sldIdLst/p:sldId", NS):
        rid = sld_id.attrib.get(f"{{{R_NS}}}id")
        if not rid or rid not in rels:
            continue
        rel_type, target = rels[rid]
        if not rel_type.endswith("/slide"):
            continue
        slide_parts.append(_resolve_target("ppt", target))

    return slide_parts or _fallback_slide_parts(pkg)


def _fallback_slide_parts(pkg: OOXMLPackage) -> list[str]:
    return sorted(
        [p for p in pkg.list_parts() if p.startswith("ppt/slides/slide") and p.endswith(".xml")]
    )


def _read_rels(pkg: OOXMLPackage, rels_part: str) -> dict[str, tuple[str, str]]:
    if not pkg.has_part(rels_part):
        return {}
    root = ET.fromstring(pkg.read_part(rels_part))
    rels: dict[str, tuple[str, str]] = {}
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.attrib.get("Id")
        if not rid:
            continue
        rels[rid] = (rel.attrib.get("Type", ""), rel.attrib.get("Target", ""))
    return rels


def _resolve_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(base_dir, target))
