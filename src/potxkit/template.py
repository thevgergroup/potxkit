from __future__ import annotations

import posixpath
from dataclasses import dataclass
from typing import Any

from .content_types import ensure_override
from .package import OOXMLPackage
from .rels import ensure_relationship
from .resources import load_base_template
from .storage import read_bytes, write_bytes
from .theme import Theme
from .validate import ValidationReport, validate_package

THEME_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
)
THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml"


@dataclass
class PotxTemplate:
    _package: OOXMLPackage
    _theme_path: str
    _theme: Theme | None = None

    @classmethod
    def open(
        cls, uri: str, *, fs_kwargs: dict[str, Any] | None = None
    ) -> "PotxTemplate":
        data = read_bytes(uri, fs_kwargs=fs_kwargs)
        pkg = OOXMLPackage(data)
        theme_path = _find_theme_part(pkg)
        return cls(pkg, theme_path)

    @classmethod
    def new(cls) -> "PotxTemplate":
        data = load_base_template()
        pkg = OOXMLPackage(data)
        theme_path = _find_theme_part(pkg)
        return cls(pkg, theme_path)

    @property
    def theme(self) -> Theme:
        if self._theme is None:
            self._theme = Theme.from_bytes(self._package.read_part(self._theme_path))
        return self._theme

    def save(self, uri: str, *, fs_kwargs: dict[str, Any] | None = None) -> None:
        if self._theme is not None:
            self._package.write_part(self._theme_path, self._theme.to_bytes())
        self._ensure_theme_relationship()
        ensure_override(self._package, f"/{self._theme_path}", THEME_CONTENT_TYPE)
        write_bytes(uri, self._package.save_bytes(), fs_kwargs=fs_kwargs)

    def validate(self) -> ValidationReport:
        return validate_package(self._package, self._theme_path)

    def _ensure_theme_relationship(self) -> None:
        if not self._package.has_part("ppt/presentation.xml"):
            return
        source_part = "ppt/presentation.xml"
        target = _rel_target(source_part, self._theme_path)
        ensure_relationship(self._package, source_part, THEME_REL_TYPE, target)


def _find_theme_part(pkg: OOXMLPackage) -> str:
    candidates = [
        name
        for name in pkg.list_parts()
        if name.startswith("ppt/theme/") and name.endswith(".xml")
    ]
    if not candidates:
        raise KeyError("No theme part found in package")
    if "ppt/theme/theme1.xml" in candidates:
        return "ppt/theme/theme1.xml"
    return sorted(candidates)[0]


def _rel_target(source_part: str, target_part: str) -> str:
    source_dir = posixpath.dirname(source_part)
    return posixpath.relpath(target_part, start=source_dir)
