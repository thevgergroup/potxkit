from __future__ import annotations

import xml.etree.ElementTree as ET

from .package import OOXMLPackage

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


class ContentTypes:
    def __init__(self, root: ET.Element) -> None:
        self._root = root

    @classmethod
    def from_bytes(cls, xml_bytes: bytes) -> "ContentTypes":
        return cls(ET.fromstring(xml_bytes))

    def ensure_override(self, part_name: str, content_type: str) -> bool:
        part = _normalize_part_name(part_name)
        for override in self._root.findall(f"{{{CT_NS}}}Override"):
            if override.attrib.get("PartName") == part:
                return False
        ET.SubElement(
            self._root,
            f"{{{CT_NS}}}Override",
            {"PartName": part, "ContentType": content_type},
        )
        return True

    def remove_override(self, part_name: str) -> bool:
        part = _normalize_part_name(part_name)
        removed = False
        for override in list(self._root.findall(f"{{{CT_NS}}}Override")):
            if override.attrib.get("PartName") == part:
                self._root.remove(override)
                removed = True
        return removed

    def to_bytes(self) -> bytes:
        return ET.tostring(self._root, encoding="utf-8", xml_declaration=True)


def ensure_override(pkg: OOXMLPackage, part_name: str, content_type: str) -> bool:
    if not pkg.has_part("[Content_Types].xml"):
        raise KeyError("[Content_Types].xml not found")
    ct = ContentTypes.from_bytes(pkg.read_part("[Content_Types].xml"))
    changed = ct.ensure_override(part_name, content_type)
    if changed:
        pkg.write_part("[Content_Types].xml", ct.to_bytes())
    return changed


def remove_override(pkg: OOXMLPackage, part_name: str) -> bool:
    if not pkg.has_part("[Content_Types].xml"):
        return False
    ct = ContentTypes.from_bytes(pkg.read_part("[Content_Types].xml"))
    changed = ct.remove_override(part_name)
    if changed:
        pkg.write_part("[Content_Types].xml", ct.to_bytes())
    return changed


def ensure_default(pkg: OOXMLPackage, extension: str, content_type: str) -> bool:
    if not pkg.has_part("[Content_Types].xml"):
        raise KeyError("[Content_Types].xml not found")
    ct = ContentTypes.from_bytes(pkg.read_part("[Content_Types].xml"))
    changed = _ensure_default_element(ct, extension, content_type)
    if changed:
        pkg.write_part("[Content_Types].xml", ct.to_bytes())
    return changed


def has_override(pkg: OOXMLPackage, part_name: str) -> bool:
    if not pkg.has_part("[Content_Types].xml"):
        return False
    ct = ContentTypes.from_bytes(pkg.read_part("[Content_Types].xml"))
    part = _normalize_part_name(part_name)
    for override in ct._root.findall(f"{{{CT_NS}}}Override"):
        if override.attrib.get("PartName") == part:
            return True
    return False


def _normalize_part_name(part_name: str) -> str:
    return part_name if part_name.startswith("/") else f"/{part_name}"


def _ensure_default_element(
    ct: ContentTypes, extension: str, content_type: str
) -> bool:
    ext = extension.lower().lstrip(".")
    for default in ct._root.findall(f"{{{CT_NS}}}Default"):
        if default.attrib.get("Extension") == ext:
            return False
    ET.SubElement(
        ct._root,
        f"{{{CT_NS}}}Default",
        {"Extension": ext, "ContentType": content_type},
    )
    return True
