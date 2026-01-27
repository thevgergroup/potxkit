from __future__ import annotations

import io
import zipfile

from potxkit.auto_layout import auto_layout
from potxkit.package import OOXMLPackage

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def test_auto_layout_creates_layouts() -> None:
    data = _minimal_pptx()
    pkg = OOXMLPackage(data)

    result = auto_layout(pkg, group_by=["p"], prefix="Auto")
    assert result.group_count == 2
    assert len(result.created_layouts) == 2


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

    pres_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" "
        "Target=\"slides/slide1.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" "
        "Target=\"slides/slide2.xml\"/>"
        "</Relationships>"
    )

    slide1 = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        f"<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\">"
        "<p:cSld><p:spTree>"
        "<p:sp><p:spPr><a:solidFill><a:srgbClr val=\"111111\"/></a:solidFill></p:spPr></p:sp>"
        "</p:spTree></p:cSld></p:sld>"
    )

    slide2 = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        f"<p:sld xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\">"
        "<p:cSld><p:spTree>"
        "<p:sp><p:spPr><a:solidFill><a:srgbClr val=\"222222\"/></a:solidFill></p:spPr></p:sp>"
        "</p:spTree></p:cSld></p:sld>"
    )

    master = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        f"<p:sldMaster xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\">"
        "<p:cSld><p:spTree/></p:cSld>"
        "<p:sldLayoutIdLst><p:sldLayoutId id=\"256\" r:id=\"rId1\"/></p:sldLayoutIdLst>"
        "</p:sldMaster>"
    )

    layout = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        f"<p:sldLayout xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\">"
        "<p:cSld><p:spTree/></p:cSld></p:sldLayout>"
    )

    master_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" "
        "Target=\"../slideLayouts/slideLayout1.xml\"/>"
        "</Relationships>"
    )

    slide_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" "
        "Target=\"../slideLayouts/slideLayout1.xml\"/>"
        "</Relationships>"
    )

    content_types = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/ppt/presentation.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>"
        "<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>"
        "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>"
        "</Types>"
    )

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("ppt/presentation.xml", presentation)
        zout.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
        zout.writestr("ppt/slides/slide1.xml", slide1)
        zout.writestr("ppt/slides/slide2.xml", slide2)
        zout.writestr("ppt/slides/_rels/slide1.xml.rels", slide_rels)
        zout.writestr("ppt/slides/_rels/slide2.xml.rels", slide_rels)
        zout.writestr("ppt/slideMasters/slideMaster1.xml", master)
        zout.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels)
        zout.writestr("ppt/slideLayouts/slideLayout1.xml", layout)
        zout.writestr("[Content_Types].xml", content_types)
    return out.getvalue()
