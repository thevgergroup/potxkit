from __future__ import annotations

import json
from dataclasses import asdict
from typing import Any

from fastmcp import FastMCP

from .audit import audit_package
from .dump_tree import DumpTreeOptions, summarize_tree
from .dump_tree import dump_tree as _dump_tree
from .layout_ops import (
    add_layout_image_shape,
    apply_palette_to_part,
    assign_slides_to_layout,
    make_layout_from_slide,
    prune_unused_layouts,
    resolve_layout_part,
    resolve_master_part,
    set_font_family_for_part,
    set_layout_background_image,
    set_layout_text_styles_for_part,
    set_master_text_styles_for_part,
    slide_size,
    strip_colors_from_part,
    strip_fonts_from_part,
)
from .layout_ops import (
    reindex_layouts as _reindex_layouts,
)
from .normalize import normalize_slide_colors, parse_slide_numbers
from .package import OOXMLPackage
from .sanitize import sanitize_slides
from .slide_index import slide_parts_in_order
from .storage import read_bytes, write_bytes
from .template import PotxTemplate

mcp = FastMCP("potxkit")


def _parse_group_by(value: str) -> list[str]:
    if not value:
        return ["p", "l"]
    parts = [item.strip() for item in value.split(",") if item.strip()]
    valid = {"p", "b", "l"}
    for item in parts:
        if item not in valid:
            raise ValueError("group_by must be a comma list of p,b,l (e.g. p,b,l)")
    return parts


def _load_pkg(path: str) -> OOXMLPackage:
    return OOXMLPackage(read_bytes(path))


def _save_pkg(pkg: OOXMLPackage, output: str) -> str:
    write_bytes(output, pkg.save_bytes())
    return output


@mcp.tool()
def info(path: str) -> dict[str, Any]:
    """Return theme colors, fonts, and validation status for a .pptx/.potx."""
    tpl = PotxTemplate.open(path)
    colors = tpl.theme.colors.as_dict()
    fonts = tpl.theme.fonts
    major = fonts.get_major()
    minor = fonts.get_minor()
    report = tpl.validate()
    return {
        "colors": colors,
        "fonts": {
            "major": major.latin if major else None,
            "minor": minor.latin if minor else None,
        },
        "validation": {
            "ok": report.ok,
            "errors": report.errors,
            "warnings": report.warnings,
        },
    }


@mcp.tool()
def validate(path: str) -> dict[str, Any]:
    """Validate a .pptx/.potx and return errors/warnings."""
    tpl = PotxTemplate.open(path)
    report = tpl.validate()
    return {"ok": report.ok, "errors": report.errors, "warnings": report.warnings}


@mcp.tool()
def dump_theme(path: str, pretty: bool = False) -> str:
    """Dump theme colors/fonts as JSON."""
    tpl = PotxTemplate.open(path)
    payload: dict[str, Any] = dict(tpl.theme.colors.as_dict())
    fonts = tpl.theme.fonts
    major = fonts.get_major()
    minor = fonts.get_minor()
    if major:
        payload["majorFont"] = major.latin
    if minor:
        payload["minorFont"] = minor.latin
    return json.dumps(payload, indent=2 if pretty else None)


@mcp.tool()
def audit(
    path: str, slides: str | None = None, group_by: str = "p,l"
) -> dict[str, Any]:
    """Audit slides for hard-coded colors, overrides, images, and backgrounds."""
    pkg = _load_pkg(path)
    slide_numbers = parse_slide_numbers(slides) if slides else None
    report = audit_package(pkg, slide_numbers, group_by=_parse_group_by(group_by))
    return {
        "theme": report.theme,
        "masters": report.masters,
        "layouts": report.layouts,
        "slides": report.slides,
        "groups": report.groups,
    }


@mcp.tool()
def dump_tree(
    path: str,
    slides: str | None = None,
    grouped: bool = False,
    include_layout: bool = False,
    include_master: bool = False,
    include_text: bool = False,
    summary: bool = False,
) -> dict[str, Any]:
    """Dump a hierarchical view of slides; use summary=true for compact output."""
    pkg = _load_pkg(path)
    slide_numbers = parse_slide_numbers(slides) if slides else None
    if grouped and not include_layout and not include_master:
        include_layout = True
        include_master = True
    options = DumpTreeOptions(
        include_layout=include_layout,
        include_master=include_master,
        include_text=include_text,
        grouped=grouped,
    )
    payload = _dump_tree(pkg, slide_numbers=slide_numbers, options=options)
    if summary:
        return {"summary": summarize_tree(payload)}
    return payload


@mcp.tool()
def normalize(
    input_path: str,
    output: str,
    mapping: dict[str, str],
    slides: str | None = None,
) -> dict[str, Any]:
    """Replace hard-coded colors with theme scheme colors using a mapping."""
    pkg = _load_pkg(input_path)
    slide_numbers = parse_slide_numbers(slides) if slides else None
    result = normalize_slide_colors(pkg, mapping, slide_numbers)
    _save_pkg(pkg, output)
    return asdict(result)


@mcp.tool()
def set_colors(
    input_path: str | None,
    output: str,
    colors: dict[str, str],
) -> str:
    """Set theme color slots (dk1, lt1, accent1, etc.)."""
    tpl = PotxTemplate.open(input_path) if input_path else PotxTemplate.new()
    theme_colors = tpl.theme.colors
    for key, value in colors.items():
        if key in {"dark1", "dk1"}:
            theme_colors.set_dark1(value)
        elif key in {"light1", "lt1"}:
            theme_colors.set_light1(value)
        elif key in {"dark2", "dk2"}:
            theme_colors.set_dark2(value)
        elif key in {"light2", "lt2"}:
            theme_colors.set_light2(value)
        elif key.startswith("accent") and key[6:].isdigit():
            theme_colors.set_accent(int(key[6:]), value)
        elif key == "hlink":
            theme_colors.set_hyperlink(value)
        elif key == "folHlink":
            theme_colors.set_followed_hyperlink(value)
    tpl.save(output)
    return output


@mcp.tool()
def set_fonts(
    input_path: str | None,
    output: str,
    major: str | None = None,
    minor: str | None = None,
) -> str:
    """Set major/minor theme fonts."""
    tpl = PotxTemplate.open(input_path) if input_path else PotxTemplate.new()
    fonts = tpl.theme.fonts
    if major:
        fonts.set_major(major)
    if minor:
        fonts.set_minor(minor)
    tpl.save(output)
    return output


@mcp.tool()
def set_theme_names(
    input_path: str | None,
    output: str,
    theme: str | None = None,
    colors: str | None = None,
    fonts: str | None = None,
) -> str:
    """Set theme, color scheme, and font scheme names for PowerPoint UI."""
    tpl = PotxTemplate.open(input_path) if input_path else PotxTemplate.new()
    if theme:
        tpl.theme.set_name(theme)
    if colors:
        tpl.theme.set_color_scheme_name(colors)
    if fonts:
        tpl.theme.set_font_scheme_name(fonts)
    tpl.save(output)
    return output


@mcp.tool()
def make_layout(
    input_path: str,
    output: str,
    from_slide: int,
    name: str,
    master_index: int = 1,
    assign_slides: str | None = None,
) -> dict[str, Any]:
    """Create a slide layout from a slide and optionally reassign slides to it."""
    pkg = _load_pkg(input_path)
    layout_part = make_layout_from_slide(pkg, from_slide, name, master_index)
    if assign_slides:
        slides = parse_slide_numbers(assign_slides)
        assign_slides_to_layout(pkg, slides, layout_part)
    _save_pkg(pkg, output)
    return {"output": output, "layout_part": layout_part}


@mcp.tool()
def set_layout(
    input_path: str,
    output: str,
    layout: str,
    palette: dict[str, str] | None = None,
    palette_none: bool = False,
    font: str | None = None,
    fonts_none: bool = False,
) -> str:
    """Update a layout's palette or fonts (or strip overrides)."""
    pkg = _load_pkg(input_path)
    layout_part = resolve_layout_part(pkg, layout)
    if palette:
        apply_palette_to_part(pkg, layout_part, palette)
    if palette_none:
        strip_colors_from_part(pkg, layout_part)
    if font:
        set_font_family_for_part(pkg, layout_part, font)
    if fonts_none:
        strip_fonts_from_part(pkg, layout_part)
    _save_pkg(pkg, output)
    return output


@mcp.tool()
def set_master(
    input_path: str,
    output: str,
    master: str = "1",
    palette: dict[str, str] | None = None,
    palette_none: bool = False,
    font: str | None = None,
    fonts_none: bool = False,
) -> str:
    """Update a slide master palette or fonts (or strip overrides)."""
    pkg = _load_pkg(input_path)
    master_part = resolve_master_part(pkg, master)
    if palette:
        apply_palette_to_part(pkg, master_part, palette)
    if palette_none:
        strip_colors_from_part(pkg, master_part)
    if font:
        set_font_family_for_part(pkg, master_part, font)
    if fonts_none:
        strip_fonts_from_part(pkg, master_part)
    _save_pkg(pkg, output)
    return output


@mcp.tool()
def set_slide(
    input_path: str,
    output: str,
    slides: str,
    layout: str | None = None,
    palette: dict[str, str] | None = None,
    palette_none: bool = False,
    font: str | None = None,
    fonts_none: bool = False,
) -> str:
    """Update slide-level palette/fonts and optionally reassign layouts."""
    pkg = _load_pkg(input_path)
    slide_numbers = parse_slide_numbers(slides)
    parts = slide_parts_in_order(pkg)
    slide_parts = []
    for num in sorted(slide_numbers):
        if num < 1 or num > len(parts):
            raise ValueError("Slide number out of range")
        slide_parts.append(parts[num - 1])

    for slide_part in slide_parts:
        if palette:
            apply_palette_to_part(pkg, slide_part, palette)
        if palette_none:
            strip_colors_from_part(pkg, slide_part)
        if font:
            set_font_family_for_part(pkg, slide_part, font)
        if fonts_none:
            strip_fonts_from_part(pkg, slide_part)
    if layout:
        layout_part = resolve_layout_part(pkg, layout)
        assign_slides_to_layout(pkg, slide_numbers, layout_part)

    _save_pkg(pkg, output)
    return output


@mcp.tool()
def set_text_styles(
    input_path: str,
    output: str,
    layout: str | None = None,
    master: str | None = None,
    title_size: float | None = None,
    body_size: float | None = None,
    title_bold: bool | None = None,
    body_bold: bool | None = None,
) -> str:
    """Set title/body text sizes and bold styles on layouts or masters."""
    pkg = _load_pkg(input_path)
    if layout:
        layout_part = resolve_layout_part(pkg, layout)
        set_layout_text_styles_for_part(
            pkg, layout_part, title_size, title_bold, body_size, body_bold
        )
    if master:
        master_part = resolve_master_part(pkg, master)
        set_master_text_styles_for_part(
            pkg, master_part, title_size, title_bold, body_size, body_bold
        )
    _save_pkg(pkg, output)
    return output


@mcp.tool()
def set_layout_bg(input_path: str, output: str, layout: str, image: str) -> str:
    """Set a layout background image."""
    pkg = _load_pkg(input_path)
    layout_part = resolve_layout_part(pkg, layout)
    set_layout_background_image(pkg, layout_part, image)
    _save_pkg(pkg, output)
    return output


@mcp.tool()
def set_layout_image(
    input_path: str,
    output: str,
    layout: str,
    image: str,
    x: float | None = None,
    y: float | None = None,
    w: float | None = None,
    h: float | None = None,
    units: str = "in",
    name: str | None = None,
) -> str:
    """Add an image layer to a layout (x/y/w/h in inches unless units=emu)."""
    pkg = _load_pkg(input_path)
    layout_part = resolve_layout_part(pkg, layout)
    cx, cy = slide_size(pkg)
    if units == "emu":
        x_emu = int(x or 0)
        y_emu = int(y or 0)
        w_emu = int(w if w is not None else cx)
        h_emu = int(h if h is not None else cy)
    else:
        factor = 914400
        x_emu = int((x or 0) * factor)
        y_emu = int((y or 0) * factor)
        w_emu = int((w if w is not None else (cx / factor)) * factor)
        h_emu = int((h if h is not None else (cy / factor)) * factor)
    add_layout_image_shape(pkg, layout_part, image, x_emu, y_emu, w_emu, h_emu, name)
    _save_pkg(pkg, output)
    return output


@mcp.tool()
def auto_layout(
    input_path: str,
    output: str,
    group_by: str = "p,l",
    prefix: str = "Auto Layout",
    master_index: int = 1,
    assign: bool = True,
    strip_colors: bool = False,
    strip_fonts: bool = False,
    palette: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Auto-generate layouts by grouping slides and optionally strip overrides."""
    pkg = _load_pkg(input_path)
    from .auto_layout import auto_layout as _auto

    result = _auto(
        pkg,
        group_by=_parse_group_by(group_by),
        prefix=prefix,
        master_index=master_index,
        assign=assign,
        strip_colors=strip_colors,
        strip_fonts=strip_fonts,
        palette=palette,
    )
    _save_pkg(pkg, output)
    return {"output": output, "layouts_created": len(result.created_layouts)}


@mcp.tool()
def prune_layouts(input_path: str, output: str) -> dict[str, Any]:
    """Remove unused slide layouts and update master references."""
    pkg = _load_pkg(input_path)
    result = prune_unused_layouts(pkg)
    _save_pkg(pkg, output)
    return {"output": output, "removed": len(result.removed_layouts)}


@mcp.tool()
def reindex_layouts(input_path: str, output: str) -> dict[str, Any]:
    """Renumber layouts and update slide/master references."""
    pkg = _load_pkg(input_path)
    result = _reindex_layouts(pkg)
    _save_pkg(pkg, output)
    return {"output": output, "layout_mapping": result.layout_mapping}


@mcp.tool()
def sanitize(input_path: str, output: str, slides: str | None = None) -> dict[str, Any]:
    """Insert missing clrMapOvr, lstStyle, and bg/noFill to avoid repair prompts."""
    pkg = _load_pkg(input_path)
    slide_numbers = parse_slide_numbers(slides) if slides else None
    result = sanitize_slides(pkg, slide_numbers)
    _save_pkg(pkg, output)
    return asdict(result)


def main() -> None:
    """Run the potxkit MCP server over stdio."""
    mcp.run()
