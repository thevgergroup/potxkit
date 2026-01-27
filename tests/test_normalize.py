from __future__ import annotations

import io
import zipfile
import xml.etree.ElementTree as ET

from potxkit.normalize import normalize_slide_colors, parse_slide_numbers
from potxkit.package import OOXMLPackage

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def test_parse_slide_numbers() -> None:
    assert parse_slide_numbers("1,3-5,8") == {1, 3, 4, 5, 8}
    assert parse_slide_numbers("5-3") == {3, 4, 5}


def test_normalize_selected_slides() -> None:
    data = _minimal_pptx()
    pkg = OOXMLPackage(data)

    result = normalize_slide_colors(pkg, {"0D0D14": "dk1"}, {1})
    assert result.replacements == 1
    assert result.slides_touched == 1

    slide1 = ET.fromstring(pkg.read_part("ppt/slides/slide1.xml"))
    slide2 = ET.fromstring(pkg.read_part("ppt/slides/slide2.xml"))

    assert slide1.find(f".//{{{A_NS}}}schemeClr") is not None
    assert slide2.find(f".//{{{A_NS}}}srgbClr") is not None


def _minimal_pptx() -> bytes:
    presentation = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        f"<p:presentation xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\">"
        "<p:sldIdLst>"
        "<p:sldId id=\"256\" r:id=\"rId1\"/>"
        "<p:sldId id=\"257\" r:id=\"rId2\"/>"
        "</p:sldIdLst>"
        "</p:presentation>"
    )

    rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" "
        "Target=\"slides/slide1.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" "
        "Target=\"slides/slide2.xml\"/>"
        "</Relationships>"
    )

    slide_template = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        f"<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\">"
        "<p:cSld>"
        "<p:spTree>"
        "<p:sp>"
        "<p:spPr>"
        "<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>"
        "</p:spPr>"
        "</p:sp>"
        "</p:spTree>"
        "</p:cSld>"
        "</p:sld>"
    )

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("ppt/presentation.xml", presentation)
        zout.writestr("ppt/_rels/presentation.xml.rels", rels)
        zout.writestr("ppt/slides/slide1.xml", slide_template.format(color="0D0D14"))
        zout.writestr("ppt/slides/slide2.xml", slide_template.format(color="FFFFFF"))
    return out.getvalue()
