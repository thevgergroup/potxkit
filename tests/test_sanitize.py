from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile

from potxkit.package import OOXMLPackage
from potxkit.sanitize import sanitize_slides


def _build_slide_pkg() -> bytes:
    slide_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        "<p:cSld>"
        "<p:bg><p:bgPr><a:effectLst/></p:bgPr></p:bg>"
        "<p:spTree>"
        "<p:sp>"
        "<p:txBody>"
        "<a:bodyPr/>"
        "<a:p><a:r><a:t>Hello</a:t></a:r></a:p>"
        "</p:txBody>"
        "</p:sp>"
        "</p:spTree>"
        "</p:cSld>"
        "</p:sld>"
    )
    presentation_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>'
        "</p:presentation>"
    )
    presentation_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        'Target="slides/slide1.xml"/>'
        "</Relationships>"
    )
    slide_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        "</Types>"
    )

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("ppt/presentation.xml", presentation_xml)
        zout.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
        zout.writestr("ppt/slides/slide1.xml", slide_xml)
        zout.writestr("ppt/slides/_rels/slide1.xml.rels", slide_rels)
        zout.writestr("[Content_Types].xml", content_types)
    return out.getvalue()


def test_sanitize_adds_defaults() -> None:
    pkg = OOXMLPackage(_build_slide_pkg())
    result = sanitize_slides(pkg)
    assert result.slides_updated == 1
    assert result.clrmap_added == 1
    assert result.lststyle_added == 1
    assert result.bg_nofill_added == 1

    root = ET.fromstring(pkg.read_part("ppt/slides/slide1.xml"))
    ns = {
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    assert root.find("p:clrMapOvr", ns) is not None
    assert root.find(".//p:txBody/a:lstStyle", ns) is not None
    assert root.find("p:cSld/p:bg/p:bgPr/a:noFill", ns) is not None
