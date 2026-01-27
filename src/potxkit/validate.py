from __future__ import annotations

import posixpath
from dataclasses import dataclass, field

from .content_types import has_override
from .package import OOXMLPackage
from .rels import parse_relationships, source_part_for


@dataclass
class ValidationReport:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors


def validate_package(pkg: OOXMLPackage, theme_path: str) -> ValidationReport:
    report = ValidationReport()
    if not pkg.has_part(theme_path):
        report.errors.append(f"Missing theme part: {theme_path}")

    if not pkg.has_part("[Content_Types].xml"):
        report.errors.append("Missing [Content_Types].xml")
    elif not has_override(pkg, theme_path):
        report.warnings.append(f"No content type override for /{theme_path}")

    _validate_relationship_targets(pkg, report)
    return report


def _validate_relationship_targets(pkg: OOXMLPackage, report: ValidationReport) -> None:
    parts = pkg.list_parts()
    rels_parts = [name for name in parts if name.endswith(".rels")]
    for rels_part in rels_parts:
        xml_bytes = pkg.read_part(rels_part)
        for rel in parse_relationships(xml_bytes):
            if rel.target_mode == "External":
                continue
            source_part = source_part_for(rels_part)
            target = _resolve_target(source_part, rel.target)
            if target and not pkg.has_part(target):
                report.errors.append(f"Missing rel target: {rels_part} -> {rel.target}")


def _resolve_target(source_part: str, target: str) -> str | None:
    if target.startswith("/"):
        return target[1:]
    if source_part == "":
        return target
    source_dir = posixpath.dirname(source_part)
    resolved = posixpath.normpath(posixpath.join(source_dir, target))
    return resolved
