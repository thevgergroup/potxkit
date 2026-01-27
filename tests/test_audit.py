from __future__ import annotations

import io
import zipfile

from potxkit.audit import audit_package
from potxkit.package import OOXMLPackage

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def test_audit_selected_slides() -> None:
    data = _minimal_pptx()
    pkg = OOXMLPackage(data)

    report = audit_package(pkg, {2})
    assert report.slides_total == 2
    assert report.slides_audited == 1
    assert report.layouts == {}
    assert report.masters == {}
    assert report.groups
    assert report.group_by == ["p", "l"]

    slide = report.per_slide[2]
    assert slide["color_counts"]["scheme"] == 1
    assert slide["has_clrMapOvr"] is True


def _minimal_pptx() -> bytes:
    presentation = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}">'
        "<p:sldIdLst>"
        '<p:sldId id="256" r:id="rId1"/>'
        '<p:sldId id="257" r:id="rId2"/>'
        "</p:sldIdLst>"
        "</p:presentation>"
    )

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        'Target="slides/slide1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        'Target="slides/slide2.xml"/>'
        "</Relationships>"
    )

    slide1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        "<p:cSld><p:spTree>"
        '<p:sp><p:spPr><a:solidFill><a:srgbClr val="0D0D14"/></a:solidFill></p:spPr></p:sp>'
        "</p:spTree></p:cSld></p:sld>"
    )

    slide2 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        "<p:cSld><p:spTree>"
        '<p:sp><p:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>'
        "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
        "</p:sld>"
    )

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("ppt/presentation.xml", presentation)
        zout.writestr("ppt/_rels/presentation.xml.rels", rels)
        zout.writestr("ppt/slides/slide1.xml", slide1)
        zout.writestr("ppt/slides/slide2.xml", slide2)
    return out.getvalue()
