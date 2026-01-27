from __future__ import annotations

import xml.etree.ElementTree as ET

from potxkit.formatting import (
    apply_color_mapping,
    set_text_font_family,
    strip_hardcoded_colors,
    strip_inline_formatting,
)

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def test_apply_color_mapping() -> None:
    root = ET.fromstring(
        f"<a:solidFill xmlns:a=\"{A_NS}\"><a:srgbClr val=\"FF0000\"/></a:solidFill>"
    )
    assert apply_color_mapping(root, {"FF0000": "accent1"}) == 1
    assert root.find(f".//{{{A_NS}}}schemeClr") is not None


def test_strip_hardcoded_colors() -> None:
    root = ET.fromstring(
        f"<a:solidFill xmlns:a=\"{A_NS}\"><a:srgbClr val=\"FF0000\"/></a:solidFill>"
    )
    assert strip_hardcoded_colors(root) == 1
    assert root.find(f".//{{{A_NS}}}srgbClr") is None


def test_strip_inline_formatting() -> None:
    root = ET.fromstring(
        f"<a:p xmlns:a=\"{A_NS}\"><a:r><a:rPr><a:latin typeface=\"X\"/></a:rPr></a:r></a:p>"
    )
    assert strip_inline_formatting(root) == 1
    assert root.find(f".//{{{A_NS}}}rPr") is None


def test_set_text_font_family() -> None:
    root = ET.fromstring(
        f"<a:p xmlns:a=\"{A_NS}\"><a:r><a:rPr/></a:r></a:p>"
    )
    assert set_text_font_family(root, "Aptos") == 1
    latin = root.find(f".//{{{A_NS}}}latin")
    assert latin is not None
    assert latin.attrib.get("typeface") == "Aptos"
