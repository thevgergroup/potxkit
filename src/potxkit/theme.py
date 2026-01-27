from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": A_NS}
ET.register_namespace("a", A_NS)


@dataclass
class ThemeFontSpec:
    latin: str
    east_asian: str | None = None
    complex_script: str | None = None


class ThemeColors:
    def __init__(self, clr_scheme: ET.Element) -> None:
        self._clr_scheme = clr_scheme

    def get_dark1(self) -> str | None:
        return self._get_slot_hex("dk1")

    def get_light1(self) -> str | None:
        return self._get_slot_hex("lt1")

    def get_dark2(self) -> str | None:
        return self._get_slot_hex("dk2")

    def get_light2(self) -> str | None:
        return self._get_slot_hex("lt2")

    def get_accent(self, index: int) -> str | None:
        return self._get_slot_hex(f"accent{index}")

    def get_hyperlink(self) -> str | None:
        return self._get_slot_hex("hlink")

    def get_followed_hyperlink(self) -> str | None:
        return self._get_slot_hex("folHlink")

    def set_dark1(self, value: str) -> None:
        self._set_slot_hex("dk1", value)

    def set_light1(self, value: str) -> None:
        self._set_slot_hex("lt1", value)

    def set_dark2(self, value: str) -> None:
        self._set_slot_hex("dk2", value)

    def set_light2(self, value: str) -> None:
        self._set_slot_hex("lt2", value)

    def set_accent(self, index: int, value: str) -> None:
        self._set_slot_hex(f"accent{index}", value)

    def set_hyperlink(self, value: str) -> None:
        self._set_slot_hex("hlink", value)

    def set_followed_hyperlink(self, value: str) -> None:
        self._set_slot_hex("folHlink", value)

    def as_dict(self) -> dict[str, str | None]:
        slots = [
            "dk1",
            "lt1",
            "dk2",
            "lt2",
            "accent1",
            "accent2",
            "accent3",
            "accent4",
            "accent5",
            "accent6",
            "hlink",
            "folHlink",
        ]
        return {slot: self._get_slot_hex(slot) for slot in slots}

    def _get_slot_hex(self, name: str) -> str | None:
        slot = self._clr_scheme.find(f"a:{name}", NS)
        if slot is None:
            return None
        srgb = slot.find("a:srgbClr", NS)
        if srgb is not None and srgb.attrib.get("val"):
            return f"#{srgb.attrib['val'].upper()}"
        sysclr = slot.find("a:sysClr", NS)
        if sysclr is not None and sysclr.attrib.get("lastClr"):
            return f"#{sysclr.attrib['lastClr'].upper()}"
        return None

    def _set_slot_hex(self, name: str, value: str) -> None:
        slot = self._clr_scheme.find(f"a:{name}", NS)
        if slot is None:
            slot = ET.SubElement(self._clr_scheme, f"{{{A_NS}}}{name}")
        for child in list(slot):
            slot.remove(child)
        srgb = ET.SubElement(slot, f"{{{A_NS}}}srgbClr")
        srgb.set("val", _normalize_hex(value))


class ThemeFonts:
    def __init__(self, font_scheme: ET.Element) -> None:
        self._font_scheme = font_scheme

    def get_major(self) -> ThemeFontSpec | None:
        return self._get_font_spec("majorFont")

    def get_minor(self) -> ThemeFontSpec | None:
        return self._get_font_spec("minorFont")

    def set_major(
        self,
        latin: str,
        east_asian: str | None = None,
        complex_script: str | None = None,
    ) -> None:
        self._set_font_spec("majorFont", latin, east_asian, complex_script)

    def set_minor(
        self,
        latin: str,
        east_asian: str | None = None,
        complex_script: str | None = None,
    ) -> None:
        self._set_font_spec("minorFont", latin, east_asian, complex_script)

    def _get_font_spec(self, name: str) -> ThemeFontSpec | None:
        node = self._font_scheme.find(f"a:{name}", NS)
        if node is None:
            return None
        latin = node.find("a:latin", NS)
        if latin is None:
            return None
        ea = node.find("a:ea", NS)
        cs = node.find("a:cs", NS)
        return ThemeFontSpec(
            latin=latin.attrib.get("typeface", ""),
            east_asian=ea.attrib.get("typeface") if ea is not None else None,
            complex_script=cs.attrib.get("typeface") if cs is not None else None,
        )

    def _set_font_spec(
        self,
        name: str,
        latin: str,
        east_asian: str | None,
        complex_script: str | None,
    ) -> None:
        node = self._font_scheme.find(f"a:{name}", NS)
        if node is None:
            node = ET.SubElement(self._font_scheme, f"{{{A_NS}}}{name}")
        _set_font_child(node, "latin", latin)
        if east_asian is not None:
            _set_font_child(node, "ea", east_asian)
        if complex_script is not None:
            _set_font_child(node, "cs", complex_script)


class Theme:
    def __init__(self, root: ET.Element) -> None:
        self._root = root
        theme_elements = root.find("a:themeElements", NS)
        if theme_elements is None:
            raise ValueError("Theme is missing themeElements")
        clr_scheme = theme_elements.find("a:clrScheme", NS)
        font_scheme = theme_elements.find("a:fontScheme", NS)
        if clr_scheme is None:
            raise ValueError("Theme is missing clrScheme")
        if font_scheme is None:
            raise ValueError("Theme is missing fontScheme")
        self._clr_scheme = clr_scheme
        self._font_scheme = font_scheme
        self.colors = ThemeColors(clr_scheme)
        self.fonts = ThemeFonts(font_scheme)

    @classmethod
    def from_bytes(cls, xml_bytes: bytes) -> "Theme":
        root = ET.fromstring(xml_bytes)
        return cls(root)

    def to_bytes(self) -> bytes:
        return ET.tostring(self._root, encoding="utf-8", xml_declaration=True)

    def get_name(self) -> str | None:
        return self._root.attrib.get("name")

    def set_name(self, value: str) -> None:
        self._root.set("name", value)

    def get_color_scheme_name(self) -> str | None:
        return self._clr_scheme.attrib.get("name")

    def set_color_scheme_name(self, value: str) -> None:
        self._clr_scheme.set("name", value)

    def get_font_scheme_name(self) -> str | None:
        return self._font_scheme.attrib.get("name")

    def set_font_scheme_name(self, value: str) -> None:
        self._font_scheme.set("name", value)


def _normalize_hex(value: str) -> str:
    value = value.strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", value):
        raise ValueError(f"Invalid hex color: {value}")
    return value.upper()


def _set_font_child(node: ET.Element, tag: str, typeface: str) -> None:
    child = node.find(f"a:{tag}", NS)
    if child is None:
        child = ET.SubElement(node, f"{{{A_NS}}}{tag}")
    child.set("typeface", typeface)
