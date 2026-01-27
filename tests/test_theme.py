from __future__ import annotations

from pathlib import Path

from potxkit import PotxTemplate
from potxkit.rels import parse_relationships

from conftest import build_minimal_potx


def test_theme_edit_roundtrip(tmp_path: Path) -> None:
    base = tmp_path / "base.potx"
    base.write_bytes(build_minimal_potx())

    tpl = PotxTemplate.open(str(base))
    tpl.theme.colors.set_accent(1, "#112233")
    tpl.theme.fonts.set_major("Aptos Display")

    out = tmp_path / "out.potx"
    tpl.save(str(out))

    reopened = PotxTemplate.open(str(out))
    assert reopened.theme.colors.get_accent(1) == "#112233"
    assert reopened.theme.fonts.get_major().latin == "Aptos Display"

    report = reopened.validate()
    assert report.ok


def test_save_adds_theme_relationship(tmp_path: Path) -> None:
    base = tmp_path / "base.potx"
    base.write_bytes(build_minimal_potx(include_theme_rel=False))

    tpl = PotxTemplate.open(str(base))
    out = tmp_path / "out.potx"
    tpl.save(str(out))

    rels_path = "ppt/_rels/presentation.xml.rels"
    rels_data = PotxTemplate.open(str(out))._package.read_part(rels_path)
    relationships = parse_relationships(rels_data)
    assert any(
        rel.type
        == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        for rel in relationships
    )


def test_new_template_roundtrip(tmp_path: Path) -> None:
    tpl = PotxTemplate.new()
    tpl.theme.colors.set_accent(2, "#123456")

    out = tmp_path / "new.potx"
    tpl.save(str(out))

    reopened = PotxTemplate.open(str(out))
    assert reopened.theme.colors.get_accent(2) == "#123456"
