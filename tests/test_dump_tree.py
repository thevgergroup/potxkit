from __future__ import annotations

import io
import json
import zipfile

from potxkit.dump_tree import DumpTreeOptions, dump_tree, summarize_tree
from potxkit.package import OOXMLPackage


def _build_deck() -> bytes:
    presentation_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/></p:sldIdLst>"
        "</p:presentation>"
    )
    presentation_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" "
        "Target=\"slides/slide1.xml\"/>"
        "</Relationships>"
    )
    slide_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<p:cSld>"
        "<p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></p:bgPr></p:bg>"
        "<p:spTree>"
        "<p:sp>"
        "<p:nvSpPr><p:cNvPr id=\"2\" name=\"Title\"/></p:nvSpPr>"
        "<p:spPr><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></p:spPr>"
        "<p:txBody>"
        "<a:bodyPr/>"
        "<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:rPr>"
        "<a:t>Hello</a:t></a:r></a:p>"
        "</p:txBody>"
        "</p:sp>"
        "</p:spTree>"
        "</p:cSld>"
        "</p:sld>"
    )
    layout_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
        "<p:cSld><p:spTree/></p:cSld>"
        "</p:sldLayout>"
    )
    layout_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" "
        "Target=\"../slideMasters/slideMaster1.xml\"/>"
        "</Relationships>"
    )
    slide_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" "
        "Target=\"../slideLayouts/slideLayout1.xml\"/>"
        "</Relationships>"
    )
    master_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
        "<p:cSld><p:spTree/></p:cSld>"
        "</p:sldMaster>"
    )
    master_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" "
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
        "<Override PartName=\"/ppt/slides/slide1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
        "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>"
        "<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>"
        "</Types>"
    )

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("ppt/presentation.xml", presentation_xml)
        zout.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
        zout.writestr("ppt/slides/slide1.xml", slide_xml)
        zout.writestr("ppt/slides/_rels/slide1.xml.rels", slide_rels)
        zout.writestr("ppt/slideLayouts/slideLayout1.xml", layout_xml)
        zout.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels)
        zout.writestr("ppt/slideMasters/slideMaster1.xml", master_xml)
        zout.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels)
        zout.writestr("[Content_Types].xml", content_types)
    return out.getvalue()


def test_dump_tree_includes_background_and_shapes() -> None:
    pkg = OOXMLPackage(_build_deck())
    payload = dump_tree(
        pkg,
        options=DumpTreeOptions(
            include_layout=True, include_master=True, include_text=True, grouped=True
        ),
    )

    slides = payload["slides"]
    assert len(slides) == 1
    slide = slides[0]
    assert slide["local"]["background"]["fill"]["type"] == "solid"
    assert slide["local"]["background"]["fill"]["color"]["value"] == "FF0000"

    shapes = slide["local"]["shapes"]
    assert shapes[0]["type"] == "shape"
    assert shapes[0]["fill"]["type"] == "solid"
    assert shapes[0]["fill"]["color"]["value"] == "accent1"
    assert shapes[0]["text"]["colors"][0]["value"] == "00FF00"

    assert slide["slideLayout"]["part"] == "ppt/slideLayouts/slideLayout1.xml"
    assert slide["slideMaster"]["part"] == "ppt/slideMasters/slideMaster1.xml"


def test_dump_tree_summary_includes_local() -> None:
    pkg = OOXMLPackage(_build_deck())
    payload = dump_tree(
        pkg,
        options=DumpTreeOptions(
            include_layout=True, include_master=True, include_text=True, grouped=True
        ),
    )
    lines = summarize_tree(payload)
    assert any(line.startswith("slide 1:") for line in lines)
    assert any("local:" in line for line in lines)
