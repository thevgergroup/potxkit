from __future__ import annotations

import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

from pptx import Presentation

from potxkit.package import OOXMLPackage

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PRESENTATION_PART = "/ppt/presentation.xml"
TEMPLATE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
)
ET.register_namespace("p", P_NS)
ET.register_namespace("a", A_NS)


def main() -> None:
    out_path = Path("src/potxkit/data/base.potx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_pptx = Path(tmp_dir) / "base.pptx"
        Presentation().save(tmp_pptx)
        data = tmp_pptx.read_bytes()

    pkg = OOXMLPackage(data)
    pkg.write_part(
        "[Content_Types].xml",
        _update_content_types(pkg.read_part("[Content_Types].xml")),
    )
    pkg.write_part(
        "ppt/slideMasters/slideMaster1.xml",
        _apply_master_background(pkg.read_part("ppt/slideMasters/slideMaster1.xml")),
    )
    out_path.write_bytes(pkg.save_bytes())


def _update_content_types(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    for override in root.findall(f"{{{CT_NS}}}Override"):
        if override.attrib.get("PartName") == PRESENTATION_PART:
            override.set("ContentType", TEMPLATE_CONTENT_TYPE)
            break
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _apply_master_background(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    c_sld = root.find(f"{{{P_NS}}}cSld")
    if c_sld is None:
        return xml_bytes

    bg = c_sld.find(f"{{{P_NS}}}bg")
    if bg is None:
        bg = ET.SubElement(c_sld, f"{{{P_NS}}}bg")
    for child in list(bg):
        bg.remove(child)

    bg_pr = ET.SubElement(bg, f"{{{P_NS}}}bgPr")
    grad = ET.SubElement(bg_pr, f"{{{A_NS}}}gradFill", {"rotWithShape": "1"})
    gs_list = ET.SubElement(grad, f"{{{A_NS}}}gsLst")

    stops = [
        (0, "0B0B0E"),
        (55000, "0B2E66"),
        (100000, "4B0F6B"),
    ]
    for pos, color in stops:
        gs = ET.SubElement(gs_list, f"{{{A_NS}}}gs", {"pos": str(pos)})
        ET.SubElement(gs, f"{{{A_NS}}}srgbClr", {"val": color})

    ET.SubElement(grad, f"{{{A_NS}}}lin", {"ang": "5400000", "scaled": "0"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


if __name__ == "__main__":
    main()
