from __future__ import annotations

import xml.etree.ElementTree as ET

from potxkit.typography import (
    detect_placeholder_styles,
    set_layout_text_styles,
    set_master_text_styles,
)

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def test_detect_placeholder_styles() -> None:
    xml = (
        f'<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        "<p:cSld><p:spTree>"
        '<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
        '<p:txBody><a:p><a:r><a:rPr sz="3000" b="1"/></a:r></a:p></p:txBody>'
        "</p:sp>"
        "</p:spTree></p:cSld></p:sld>"
    )
    root = ET.fromstring(xml)
    styles = detect_placeholder_styles(root)
    assert styles["title"]["size_pt"] == 30
    assert styles["title"]["bold"] is True


def test_set_layout_text_styles() -> None:
    xml = (
        f'<p:sldLayout xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        "<p:cSld><p:spTree>"
        '<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
        "<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody></p:sp>"
        "</p:spTree></p:cSld></p:sldLayout>"
    )
    root = ET.fromstring(xml)
    assert set_layout_text_styles(root, 28, True, None, None) > 0
    def_rpr = root.find(f".//{{{A_NS}}}defRPr")
    assert def_rpr is not None
    assert def_rpr.attrib.get("sz") == "2800"
    assert def_rpr.attrib.get("b") == "1"


def test_set_master_text_styles() -> None:
    xml = (
        f'<p:sldMaster xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        "<p:txStyles><p:titleStyle><a:lvl1pPr/></p:titleStyle></p:txStyles>"
        "</p:sldMaster>"
    )
    root = ET.fromstring(xml)
    assert set_master_text_styles(root, 32, False, None, None) > 0
    def_rpr = root.find(f".//{{{A_NS}}}defRPr")
    assert def_rpr is not None
    assert def_rpr.attrib.get("sz") == "3200"
    assert def_rpr.attrib.get("b") == "0"
