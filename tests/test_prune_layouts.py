from __future__ import annotations

import io
import zipfile
import xml.etree.ElementTree as ET

from potxkit.layout_ops import prune_unused_layouts
from potxkit.package import OOXMLPackage
from potxkit.rels import parse_relationships


def _build_minimal_deck() -> bytes:
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
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
        "<p:cSld><p:spTree/></p:cSld>"
        "</p:sld>"
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
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<p:sldLayoutIdLst>"
        "<p:sldLayoutId id=\"256\" r:id=\"rId1\"/>"
        "<p:sldLayoutId id=\"257\" r:id=\"rId2\"/>"
        "</p:sldLayoutIdLst>"
        "</p:sldMaster>"
    )
    master_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" "
        "Target=\"../slideLayouts/slideLayout1.xml\"/>"
        "<Relationship Id=\"rId2\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" "
        "Target=\"../slideLayouts/slideLayout2.xml\"/>"
        "</Relationships>"
    )
    layout_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" "
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
        "<p:cSld><p:spTree/></p:cSld>"
        "</p:sldLayout>"
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
        "<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>"
        "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>"
        "<Override PartName=\"/ppt/slideLayouts/slideLayout2.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>"
        "</Types>"
    )

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("ppt/presentation.xml", presentation_xml)
        zout.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
        zout.writestr("ppt/slides/slide1.xml", slide_xml)
        zout.writestr("ppt/slides/_rels/slide1.xml.rels", slide_rels)
        zout.writestr("ppt/slideMasters/slideMaster1.xml", master_xml)
        zout.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels)
        zout.writestr("ppt/slideLayouts/slideLayout1.xml", layout_xml)
        zout.writestr("ppt/slideLayouts/slideLayout2.xml", layout_xml)
        zout.writestr("[Content_Types].xml", content_types)
    return out.getvalue()


def test_prune_unused_layouts_removes_layout_and_master_refs() -> None:
    pkg = OOXMLPackage(_build_minimal_deck())

    result = prune_unused_layouts(pkg)
    assert "ppt/slideLayouts/slideLayout2.xml" in result.removed_layouts
    assert pkg.has_part("ppt/slideLayouts/slideLayout1.xml")
    assert not pkg.has_part("ppt/slideLayouts/slideLayout2.xml")

    rels = parse_relationships(
        pkg.read_part("ppt/slideMasters/_rels/slideMaster1.xml.rels")
    )
    targets = {rel.target for rel in rels}
    assert "../slideLayouts/slideLayout2.xml" not in targets

    master_root = ET.fromstring(pkg.read_part("ppt/slideMasters/slideMaster1.xml"))
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    layout_ids = [
        node.attrib.get(f"{{{r_ns}}}id")
        for node in master_root.findall(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}sldLayoutId"
        )
    ]
    assert "rId2" not in layout_ids

    ct_root = ET.fromstring(pkg.read_part("[Content_Types].xml"))
    overrides = [
        node.attrib.get("PartName")
        for node in ct_root.findall(
            "{http://schemas.openxmlformats.org/package/2006/content-types}Override"
        )
    ]
    assert "/ppt/slideLayouts/slideLayout2.xml" not in overrides
