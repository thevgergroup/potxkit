from __future__ import annotations

import argparse
import json
import sys
import xml.etree.ElementTree as ET
from typing import Any

from . import PotxTemplate
from .audit import AuditReport, audit_package
from .auto_layout import auto_layout
from .layout_ops import (
    apply_palette_to_part,
    assign_slides_to_layout,
    add_layout_image_shape,
    make_layout_from_slide,
    prune_unused_layouts,
    reindex_layouts,
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
from .normalize import NormalizeResult, normalize_slide_colors, parse_slide_numbers
from .package import OOXMLPackage
from .sanitize import sanitize_slides
from .dump_tree import DumpTreeOptions, dump_tree, summarize_tree
from .slide_index import slide_parts_in_order
from .storage import read_bytes, write_bytes
from .typography import detect_placeholder_styles


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="potxkit",
        description="Edit PowerPoint .potx themes (colors and fonts).",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    info_parser = subparsers.add_parser("info", help="Show theme info and validation")
    info_parser.add_argument("path", help="Path to .potx/.pptx")

    new_parser = subparsers.add_parser(
        "new", help="Create a new .potx from the bundled base"
    )
    new_parser.add_argument("output", help="Output .potx path")

    apply_parser = subparsers.add_parser(
        "apply-palette", help="Apply a palette JSON to a template"
    )
    apply_parser.add_argument("palette", help="Path to palette JSON")
    apply_parser.add_argument("output", help="Output .potx path")
    apply_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .potx/.pptx (uses bundled base if omitted)",
    )

    validate_parser = subparsers.add_parser(
        "validate", help="Validate theme relationships and content types"
    )
    validate_parser.add_argument("path", help="Path to .potx/.pptx")

    audit_parser = subparsers.add_parser(
        "audit", help="Report per-slide color usage and overrides"
    )
    audit_parser.add_argument("path", help="Path to .pptx/.potx")
    audit_parser.add_argument(
        "--slides",
        help="Slide numbers/ranges to audit (e.g. 1,3-5,8). Defaults to all.",
    )
    audit_parser.add_argument("--output", help="Optional JSON output path")
    audit_parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON")
    audit_parser.add_argument(
        "--summary", action="store_true", help="Print a human-readable summary"
    )
    audit_parser.add_argument(
        "--details",
        action="store_true",
        help="Include per-slide details in summary output",
    )
    audit_parser.add_argument(
        "--group-by",
        help="Group slides by palette/background/layout (e.g. p,b,l). Default: p,l",
    )

    normalize_parser = subparsers.add_parser(
        "normalize", help="Replace hard-coded colors with theme colors"
    )
    normalize_parser.add_argument("mapping", help="Path to mapping JSON")
    normalize_parser.add_argument("output", help="Output .pptx/.potx path")
    normalize_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )
    normalize_parser.add_argument(
        "--slides",
        help="Slide numbers/ranges to edit (e.g. 1,3-5,8). Defaults to all.",
    )
    normalize_parser.add_argument(
        "--report",
        help="Optional JSON report path for replacements per slide.",
    )

    palette_parser = subparsers.add_parser(
        "palette-template", help="Print or write an example palette JSON"
    )
    palette_parser.add_argument("--output", help="Optional JSON output path")
    palette_parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON")

    dump_parser = subparsers.add_parser(
        "dump-theme", help="Write the theme colors and fonts as JSON"
    )
    dump_parser.add_argument("path", help="Path to .potx/.pptx")
    dump_parser.add_argument("--output", help="Optional JSON output path")
    dump_parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON")

    colors_parser = subparsers.add_parser("set-colors", help="Set theme colors")
    colors_parser.add_argument("output", help="Output .potx path")
    colors_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .potx/.pptx (uses bundled base if omitted)",
    )
    colors_parser.add_argument("--dark1")
    colors_parser.add_argument("--light1")
    colors_parser.add_argument("--dark2")
    colors_parser.add_argument("--light2")
    for idx in range(1, 7):
        colors_parser.add_argument(f"--accent{idx}")
    colors_parser.add_argument("--hlink")
    colors_parser.add_argument("--folHlink")

    fonts_parser = subparsers.add_parser("set-fonts", help="Set theme fonts")
    fonts_parser.add_argument("output", help="Output .potx path")
    fonts_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .potx/.pptx (uses bundled base if omitted)",
    )
    fonts_parser.add_argument("--major", help="Major font (Latin)")
    fonts_parser.add_argument("--minor", help="Minor font (Latin)")

    master_parser = subparsers.add_parser(
        "set-master", help="Update a slide master's palette or fonts"
    )
    master_parser.add_argument("--master", default="1")
    master_parser.add_argument("--palette")
    master_parser.add_argument("--palette-none", action="store_true")
    master_parser.add_argument("--font")
    master_parser.add_argument("--fonts-none", action="store_true")
    master_parser.add_argument("output", help="Output .pptx/.potx path")
    master_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )

    names_parser = subparsers.add_parser(
        "set-theme-names", help="Set theme, color scheme, and font scheme names"
    )
    names_parser.add_argument("output", help="Output .potx path")
    names_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .potx/.pptx (uses bundled base if omitted)",
    )
    names_parser.add_argument("--theme", help="Theme name (slide master UI)")
    names_parser.add_argument("--colors", help="Color scheme name")
    names_parser.add_argument("--fonts", help="Font scheme name")

    make_layout_parser = subparsers.add_parser(
        "make-layout", help="Create a layout from a slide"
    )
    make_layout_parser.add_argument("--from-slide", type=int, required=True)
    make_layout_parser.add_argument("--name", required=True)
    make_layout_parser.add_argument("--master", type=int, default=1)
    make_layout_parser.add_argument(
        "--assign-slides",
        help="Slide numbers/ranges to assign to the new layout",
    )
    make_layout_parser.add_argument("output", help="Output .pptx/.potx path")
    make_layout_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )

    set_layout_parser = subparsers.add_parser(
        "set-layout", help="Update a slide layout's palette or fonts"
    )
    set_layout_parser.add_argument("--layout", required=True)
    set_layout_parser.add_argument("--palette")
    set_layout_parser.add_argument("--palette-none", action="store_true")
    set_layout_parser.add_argument("--font")
    set_layout_parser.add_argument("--fonts-none", action="store_true")
    set_layout_parser.add_argument("output", help="Output .pptx/.potx path")
    set_layout_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )

    set_slide_parser = subparsers.add_parser(
        "set-slide", help="Update a slide's palette, fonts, or layout"
    )
    set_slide_parser.add_argument(
        "--slide",
        type=int,
        help="Single slide number (1-based)",
    )
    set_slide_parser.add_argument(
        "--slides",
        help="Slide numbers/ranges (e.g. 1,3-5,8)",
    )
    set_slide_parser.add_argument("--layout")
    set_slide_parser.add_argument("--palette")
    set_slide_parser.add_argument("--palette-none", action="store_true")
    set_slide_parser.add_argument("--font")
    set_slide_parser.add_argument("--fonts-none", action="store_true")
    set_slide_parser.add_argument("output", help="Output .pptx/.potx path")
    set_slide_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )

    text_styles_parser = subparsers.add_parser(
        "set-text-styles", help="Set title/body text sizes and bold styles"
    )
    text_styles_parser.add_argument("--layout")
    text_styles_parser.add_argument("--master")
    text_styles_parser.add_argument("--from-slide", type=int)
    text_styles_parser.add_argument("--styles", help="Path to styles JSON")
    text_styles_parser.add_argument("--title-size", type=float)
    text_styles_parser.add_argument("--body-size", type=float)
    text_styles_parser.add_argument("--title-bold", action="store_true")
    text_styles_parser.add_argument("--title-regular", action="store_true")
    text_styles_parser.add_argument("--body-bold", action="store_true")
    text_styles_parser.add_argument("--body-regular", action="store_true")
    text_styles_parser.add_argument("output", help="Output .pptx/.potx path")
    text_styles_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )

    layout_bg_parser = subparsers.add_parser(
        "set-layout-bg", help="Set a layout background image"
    )
    layout_bg_parser.add_argument("--layout", required=True)
    layout_bg_parser.add_argument("--image", required=True)
    layout_bg_parser.add_argument("output", help="Output .pptx/.potx path")
    layout_bg_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )

    layout_img_parser = subparsers.add_parser(
        "set-layout-image", help="Add an image layer to a layout"
    )
    layout_img_parser.add_argument("--layout", required=True)
    layout_img_parser.add_argument("--image", required=True)
    layout_img_parser.add_argument("--x", type=float)
    layout_img_parser.add_argument("--y", type=float)
    layout_img_parser.add_argument("--w", type=float)
    layout_img_parser.add_argument("--h", type=float)
    layout_img_parser.add_argument(
        "--units",
        choices=["in", "emu"],
        default="in",
        help="Units for x/y/w/h (default: in)",
    )
    layout_img_parser.add_argument("--name")
    layout_img_parser.add_argument("output", help="Output .pptx/.potx path")
    layout_img_parser.add_argument(
        "--input",
        dest="input_path",
        help="Optional input .pptx/.potx (uses bundled base if omitted)",
    )

    prune_layouts_parser = subparsers.add_parser(
        "prune-layouts", help="Remove unused slide layouts"
    )
    prune_layouts_parser.add_argument("output", help="Output .pptx/.potx path")
    prune_layouts_parser.add_argument(
        "--input",
        dest="input_path",
        required=True,
        help="Input .pptx/.potx path",
    )
    prune_layouts_parser.add_argument(
        "--keep",
        action="append",
        default=[],
        help="Layout selector to keep (index or name). Can be repeated.",
    )

    reindex_layouts_parser = subparsers.add_parser(
        "reindex-layouts", help="Renumber layouts and update references"
    )
    reindex_layouts_parser.add_argument("output", help="Output .pptx/.potx path")
    reindex_layouts_parser.add_argument(
        "--input",
        dest="input_path",
        required=True,
        help="Input .pptx/.potx path",
    )

    sanitize_parser = subparsers.add_parser(
        "sanitize", help="Add missing OOXML defaults to slides"
    )
    sanitize_parser.add_argument("output", help="Output .pptx/.potx path")
    sanitize_parser.add_argument(
        "--input",
        dest="input_path",
        required=True,
        help="Input .pptx/.potx path",
    )
    sanitize_parser.add_argument(
        "--slides", help="Comma-separated slide numbers or ranges (e.g. 1,3-5)"
    )

    dump_tree_parser = subparsers.add_parser(
        "dump-tree", help="Dump a hierarchical view of slides"
    )
    dump_tree_parser.add_argument("path", help="Input .pptx/.potx path")
    dump_tree_parser.add_argument("--slides", help="Comma-separated slide numbers or ranges")
    dump_tree_parser.add_argument("--layout", action="store_true")
    dump_tree_parser.add_argument("--master", action="store_true")
    dump_tree_parser.add_argument("--text", action="store_true")
    dump_tree_parser.add_argument(
        "--grouped",
        action="store_true",
        help="Group output by slideMaster/slideLayout/local blocks",
    )
    dump_tree_parser.add_argument("--output", help="Write JSON to a file")
    dump_tree_parser.add_argument("--pretty", action="store_true")
    dump_tree_parser.add_argument(
        "--summary",
        action="store_true",
        help="Print a human-friendly summary instead of JSON",
    )

    auto_layout_parser = subparsers.add_parser(
        "auto-layout", help="Auto-generate layouts by grouping slides"
    )
    auto_layout_parser.add_argument("output", help="Output .pptx/.potx path")
    auto_layout_parser.add_argument(
        "--input",
        dest="input_path",
        help="Input .pptx/.potx",
        required=True,
    )
    auto_layout_parser.add_argument(
        "--group-by",
        help="Group slides by palette/background/layout (e.g. p,b,l). Default: p,l",
    )
    auto_layout_parser.add_argument("--prefix", default="Auto Layout")
    auto_layout_parser.add_argument("--master", type=int, default=1)
    auto_layout_parser.add_argument("--no-assign", action="store_true")
    auto_layout_parser.add_argument("--strip-colors", action="store_true")
    auto_layout_parser.add_argument("--strip-fonts", action="store_true")
    auto_layout_parser.add_argument("--palette")

    args = parser.parse_args(argv)

    if args.command == "info":
        return _handle_info(args.path)
    if args.command == "new":
        return _handle_new(args.output)
    if args.command == "apply-palette":
        return _handle_apply_palette(args.palette, args.output, args.input_path)
    if args.command == "validate":
        return _handle_validate(args.path)
    if args.command == "audit":
        return _handle_audit(args)
    if args.command == "normalize":
        return _handle_normalize(args)
    if args.command == "palette-template":
        return _handle_palette_template(args.output, args.pretty)
    if args.command == "dump-theme":
        return _handle_dump_theme(args.path, args.output, args.pretty)
    if args.command == "set-colors":
        return _handle_set_colors(args)
    if args.command == "set-fonts":
        return _handle_set_fonts(args)
    if args.command == "set-master":
        return _handle_set_master(args)
    if args.command == "set-theme-names":
        return _handle_set_theme_names(args)
    if args.command == "make-layout":
        return _handle_make_layout(args)
    if args.command == "set-layout":
        return _handle_set_layout(args)
    if args.command == "set-slide":
        return _handle_set_slide(args)
    if args.command == "set-text-styles":
        return _handle_set_text_styles(args)
    if args.command == "set-layout-bg":
        return _handle_set_layout_bg(args)
    if args.command == "set-layout-image":
        return _handle_set_layout_image(args)
    if args.command == "prune-layouts":
        return _handle_prune_layouts(args)
    if args.command == "reindex-layouts":
        return _handle_reindex_layouts(args)
    if args.command == "sanitize":
        return _handle_sanitize(args)
    if args.command == "dump-tree":
        return _handle_dump_tree(args)
    if args.command == "auto-layout":
        return _handle_auto_layout(args)

    parser.print_help()
    return 1


def _handle_info(path: str) -> int:
    tpl = PotxTemplate.open(path)
    colors = tpl.theme.colors.as_dict()
    fonts = tpl.theme.fonts

    print("Theme colors:")
    for name, value in colors.items():
        print(f"- {name}: {value}")

    major = fonts.get_major()
    minor = fonts.get_minor()
    print("\nTheme fonts:")
    if major:
        print(f"- major: {major.latin}")
    if minor:
        print(f"- minor: {minor.latin}")

    report = tpl.validate()
    print("\nValidation:")
    print(f"- ok: {report.ok}")
    if report.errors:
        print("- errors:")
        for error in report.errors:
            print(f"  - {error}")
    if report.warnings:
        print("- warnings:")
        for warning in report.warnings:
            print(f"  - {warning}")

    return 0 if report.ok else 1


def _handle_apply_palette(palette_path: str, output: str, input_path: str | None) -> int:
    palette = _load_palette(palette_path)
    tpl = _load_template(input_path)
    _apply_palette(tpl, palette)
    tpl.save(output)
    print(f"Wrote {output}")
    return 0


def _handle_validate(path: str) -> int:
    tpl = PotxTemplate.open(path)
    report = tpl.validate()
    if report.ok:
        print("Validation OK")
        return 0

    print("Validation failed:")
    for error in report.errors:
        print(f"- {error}")
    for warning in report.warnings:
        print(f"- warning: {warning}")
    return 1


def _handle_normalize(args: argparse.Namespace) -> int:
    mapping = _load_mapping(args.mapping)
    data = read_bytes(args.input_path) if args.input_path else None
    if data is None:
        tpl = PotxTemplate.new()
        pkg = tpl._package
    else:
        pkg = OOXMLPackage(data)

    slide_numbers = parse_slide_numbers(args.slides) if args.slides else None
    result = normalize_slide_colors(pkg, mapping, slide_numbers)
    write_bytes(args.output, pkg.save_bytes())
    _print_normalize_result(result)

    if args.report:
        with open(args.report, "w", encoding="utf-8") as handle:
            json.dump(_normalize_report(result), handle, indent=2)
        print(f"Wrote {args.report}")

    return 0


def _handle_audit(args: argparse.Namespace) -> int:
    data = read_bytes(args.path)
    pkg = OOXMLPackage(data)
    slide_numbers = parse_slide_numbers(args.slides) if args.slides else None
    try:
        group_by = _parse_group_by(args.group_by)
    except ValueError as exc:
        print(str(exc))
        return 2
    report = audit_package(pkg, slide_numbers, group_by=group_by)

    payload = _audit_report(report)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2 if args.pretty else None)
        print(f"Wrote {args.output}")
    if not args.output and not args.summary:
        json.dump(payload, sys.stdout, indent=2 if args.pretty else None)
        print()
    if args.summary:
        _print_audit_summary(report, details=args.details)

    return 0


def _handle_new(output: str) -> int:
    tpl = PotxTemplate.new()
    tpl.save(output)
    print(f"Wrote {output}")
    return 0


def _handle_dump_theme(path: str, output: str | None, pretty: bool) -> int:
    tpl = PotxTemplate.open(path)
    payload = _theme_to_json(tpl)
    if output:
        with open(output, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2 if pretty else None)
        print(f"Wrote {output}")
    else:
        json.dump(payload, sys.stdout, indent=2 if pretty else None)
        print()
    return 0


def _handle_set_colors(args: argparse.Namespace) -> int:
    tpl = _load_template(args.input_path)
    colors = tpl.theme.colors
    updates = 0

    updates += _set_if_present(colors.set_dark1, vars(args), "dark1")
    updates += _set_if_present(colors.set_light1, vars(args), "light1")
    updates += _set_if_present(colors.set_dark2, vars(args), "dark2")
    updates += _set_if_present(colors.set_light2, vars(args), "light2")
    for idx in range(1, 7):
        key = f"accent{idx}"
        if key in vars(args) and getattr(args, key) is not None:
            colors.set_accent(idx, getattr(args, key))
            updates += 1
    updates += _set_if_present(colors.set_hyperlink, vars(args), "hlink")
    updates += _set_if_present(colors.set_followed_hyperlink, vars(args), "folHlink")

    if updates == 0:
        print("No colors specified. Use --accent1, --dark1, etc.")
        return 2

    tpl.save(args.output)
    print(f"Wrote {args.output}")
    return 0


def _handle_set_fonts(args: argparse.Namespace) -> int:
    if not args.major and not args.minor:
        print("No fonts specified. Use --major and/or --minor.")
        return 2

    tpl = _load_template(args.input_path)
    if args.major:
        tpl.theme.fonts.set_major(args.major)
    if args.minor:
        tpl.theme.fonts.set_minor(args.minor)

    tpl.save(args.output)
    print(f"Wrote {args.output}")
    return 0


def _handle_palette_template(output: str | None, pretty: bool) -> int:
    payload = _example_palette()
    if output:
        with open(output, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2 if pretty else None)
        print(f"Wrote {output}")
    else:
        json.dump(payload, sys.stdout, indent=2 if pretty else None)
        print()
    return 0


def _handle_set_theme_names(args: argparse.Namespace) -> int:
    if not args.theme and not args.colors and not args.fonts:
        print("No names specified. Use --theme, --colors, and/or --fonts.")
        return 2

    tpl = _load_template(args.input_path)
    if args.theme:
        tpl.theme.set_name(args.theme)
    if args.colors:
        tpl.theme.set_color_scheme_name(args.colors)
    if args.fonts:
        tpl.theme.set_font_scheme_name(args.fonts)

    tpl.save(args.output)
    print(f"Wrote {args.output}")
    return 0


def _handle_set_master(args: argparse.Namespace) -> int:
    try:
        _validate_palette_font_args(args)
    except ValueError as exc:
        print(str(exc))
        return 2
    data = read_bytes(args.input_path) if args.input_path else None
    pkg = OOXMLPackage(data) if data else PotxTemplate.new()._package

    master_part = resolve_master_part(pkg, str(args.master))
    _apply_palette_and_fonts(pkg, master_part, args)

    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    return 0


def _handle_make_layout(args: argparse.Namespace) -> int:
    data = read_bytes(args.input_path) if args.input_path else None
    pkg = OOXMLPackage(data) if data else PotxTemplate.new()._package

    new_layout = make_layout_from_slide(
        pkg,
        slide_number=args.from_slide,
        name=args.name,
        master_index=args.master,
    )
    if args.assign_slides:
        slides = parse_slide_numbers(args.assign_slides)
        assign_slides_to_layout(pkg, slides, new_layout)

    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    print(f"Layout created: {new_layout}")
    return 0


def _handle_set_layout(args: argparse.Namespace) -> int:
    try:
        _validate_palette_font_args(args)
    except ValueError as exc:
        print(str(exc))
        return 2
    data = read_bytes(args.input_path) if args.input_path else None
    pkg = OOXMLPackage(data) if data else PotxTemplate.new()._package

    layout_part = resolve_layout_part(pkg, args.layout)
    _apply_palette_and_fonts(pkg, layout_part, args)

    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    return 0


def _handle_set_slide(args: argparse.Namespace) -> int:
    try:
        _validate_palette_font_args(args)
    except ValueError as exc:
        print(str(exc))
        return 2
    if args.slide is None and not args.slides:
        print("Provide --slide or --slides.")
        return 2
    data = read_bytes(args.input_path) if args.input_path else None
    pkg = OOXMLPackage(data) if data else PotxTemplate.new()._package

    if args.slides:
        slide_numbers = parse_slide_numbers(args.slides)
    else:
        slide_numbers = parse_slide_numbers(str(args.slide))
    slide_parts = _slide_parts_for_numbers(pkg, slide_numbers)

    for slide_part in slide_parts:
        _apply_palette_and_fonts(pkg, slide_part, args)

    if args.layout:
        layout_part = resolve_layout_part(pkg, args.layout)
        assign_slides_to_layout(pkg, slide_numbers, layout_part)

    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    return 0


def _handle_set_text_styles(args: argparse.Namespace) -> int:
    if not args.layout and not args.master:
        print("Provide --layout and/or --master.")
        return 2
    if args.title_bold and args.title_regular:
        print("Use only one of --title-bold or --title-regular.")
        return 2
    if args.body_bold and args.body_regular:
        print("Use only one of --body-bold or --body-regular.")
        return 2

    data = read_bytes(args.input_path) if args.input_path else None
    pkg = OOXMLPackage(data) if data else PotxTemplate.new()._package

    title_size = args.title_size
    body_size = args.body_size
    title_bold = True if args.title_bold else False if args.title_regular else None
    body_bold = True if args.body_bold else False if args.body_regular else None

    if args.styles:
        styles = _load_styles(args.styles)
        title = styles.get("title", {})
        body = styles.get("body", {})
        title_size = title.get("size", title_size)
        body_size = body.get("size", body_size)
        if "bold" in title:
            title_bold = bool(title.get("bold"))
        if "bold" in body:
            body_bold = bool(body.get("bold"))

    if args.from_slide:
        slide_part = _slide_parts_for_numbers(pkg, {args.from_slide})[0]
        slide_root = ET.fromstring(pkg.read_part(slide_part))
        detected = detect_placeholder_styles(slide_root)
        if title_size is None and detected.get("title", {}).get("size_pt") is not None:
            title_size = detected["title"]["size_pt"]
        if body_size is None and detected.get("body", {}).get("size_pt") is not None:
            body_size = detected["body"]["size_pt"]
        if title_bold is None and detected.get("title", {}).get("bold") is not None:
            title_bold = detected["title"]["bold"]
        if body_bold is None and detected.get("body", {}).get("bold") is not None:
            body_bold = detected["body"]["bold"]

    if args.layout:
        layout_part = resolve_layout_part(pkg, args.layout)
        set_layout_text_styles_for_part(
            pkg, layout_part, title_size, title_bold, body_size, body_bold
        )
    if args.master:
        master_part = resolve_master_part(pkg, str(args.master))
        set_master_text_styles_for_part(
            pkg, master_part, title_size, title_bold, body_size, body_bold
        )

    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    return 0


def _load_styles(path: str) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("Styles JSON must be an object")
    return data


def _handle_set_layout_bg(args: argparse.Namespace) -> int:
    data = read_bytes(args.input_path) if args.input_path else None
    pkg = OOXMLPackage(data) if data else PotxTemplate.new()._package
    layout_part = resolve_layout_part(pkg, args.layout)
    set_layout_background_image(pkg, layout_part, args.image)
    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    return 0


def _handle_set_layout_image(args: argparse.Namespace) -> int:
    data = read_bytes(args.input_path) if args.input_path else None
    pkg = OOXMLPackage(data) if data else PotxTemplate.new()._package
    layout_part = resolve_layout_part(pkg, args.layout)

    cx, cy = slide_size(pkg)
    x, y, w, h = _resolve_image_box(args, cx, cy)
    add_layout_image_shape(pkg, layout_part, args.image, x, y, w, h, args.name)
    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    return 0


def _handle_prune_layouts(args: argparse.Namespace) -> int:
    data = read_bytes(args.input_path)
    pkg = OOXMLPackage(data)

    keep_layouts = set()
    for selector in args.keep:
        keep_layouts.add(resolve_layout_part(pkg, selector))

    result = prune_unused_layouts(pkg, keep_layouts=keep_layouts or None)
    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    print(f"Layouts removed: {len(result.removed_layouts)}")
    return 0


def _handle_reindex_layouts(args: argparse.Namespace) -> int:
    data = read_bytes(args.input_path)
    pkg = OOXMLPackage(data)
    result = reindex_layouts(pkg)
    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    print(f"Layouts remapped: {len(result.layout_mapping)}")
    return 0


def _handle_sanitize(args: argparse.Namespace) -> int:
    data = read_bytes(args.input_path)
    pkg = OOXMLPackage(data)
    slide_numbers = parse_slide_numbers(args.slides) if args.slides else None
    result = sanitize_slides(pkg, slide_numbers)
    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    print(
        f"Slides updated: {result.slides_updated} "
        f"(clrMapOvr={result.clrmap_added}, lstStyle={result.lststyle_added}, "
        f"bgNoFill={result.bg_nofill_added})"
    )
    return 0


def _handle_dump_tree(args: argparse.Namespace) -> int:
    data = read_bytes(args.path)
    pkg = OOXMLPackage(data)
    slide_numbers = parse_slide_numbers(args.slides) if args.slides else None
    include_layout = args.layout
    include_master = args.master
    if args.grouped and not include_layout and not include_master:
        include_layout = True
        include_master = True
    options = DumpTreeOptions(
        include_layout=include_layout,
        include_master=include_master,
        include_text=args.text,
        grouped=args.grouped,
    )
    payload = dump_tree(pkg, slide_numbers=slide_numbers, options=options)
    if args.summary:
        lines = summarize_tree(payload)
        output = "\n".join(lines)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as handle:
                handle.write(output)
                handle.write("\n")
            print(f"Wrote {args.output}")
            return 0
        print(output)
        return 0

    if args.output:
        with open(args.output, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2 if args.pretty else None)
        print(f"Wrote {args.output}")
        return 0
    json.dump(payload, sys.stdout, indent=2 if args.pretty else None)
    print()
    return 0


def _handle_auto_layout(args: argparse.Namespace) -> int:
    data = read_bytes(args.input_path)
    pkg = OOXMLPackage(data)
    try:
        group_by = _parse_group_by(args.group_by)
    except ValueError as exc:
        print(str(exc))
        return 2
    palette = _load_mapping(args.palette) if args.palette else None

    result = auto_layout(
        pkg,
        group_by=group_by,
        prefix=args.prefix,
        master_index=args.master,
        assign=not args.no_assign,
        strip_colors=args.strip_colors,
        strip_fonts=args.strip_fonts,
        palette=palette,
    )
    write_bytes(args.output, pkg.save_bytes())
    print(f"Wrote {args.output}")
    print(f"Layouts created: {len(result.created_layouts)}")
    return 0


def _load_palette(path: str) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("Palette JSON must be an object")
    return data


def _load_mapping(path: str) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("Mapping JSON must be an object")
    return data


def _apply_palette(tpl: PotxTemplate, palette: dict[str, Any]) -> None:
    colors = tpl.theme.colors

    _set_if_present(colors.set_dark1, palette, "dark1")
    _set_if_present(colors.set_light1, palette, "light1")
    _set_if_present(colors.set_dark2, palette, "dark2")
    _set_if_present(colors.set_light2, palette, "light2")

    for idx in range(1, 7):
        key = f"accent{idx}"
        if key in palette:
            colors.set_accent(idx, palette[key])

    _set_if_present(colors.set_hyperlink, palette, "hlink")
    _set_if_present(colors.set_followed_hyperlink, palette, "folHlink")

    fonts = tpl.theme.fonts
    if "majorFont" in palette:
        fonts.set_major(palette["majorFont"])
    if "minorFont" in palette:
        fonts.set_minor(palette["minorFont"])


def _theme_to_json(tpl: PotxTemplate) -> dict[str, Any]:
    colors = tpl.theme.colors.as_dict()
    fonts = tpl.theme.fonts
    major = fonts.get_major()
    minor = fonts.get_minor()
    payload: dict[str, Any] = dict(colors)
    if major:
        payload["majorFont"] = major.latin
    if minor:
        payload["minorFont"] = minor.latin
    return payload


def _apply_palette_and_fonts(pkg: OOXMLPackage, part: str, args: argparse.Namespace) -> None:
    if args.palette:
        mapping = _load_mapping(args.palette)
        apply_palette_to_part(pkg, part, mapping)
    if args.palette_none:
        strip_colors_from_part(pkg, part)
    if args.font:
        set_font_family_for_part(pkg, part, args.font)
    if args.fonts_none:
        strip_fonts_from_part(pkg, part)


def _validate_palette_font_args(args: argparse.Namespace) -> None:
    if args.palette and args.palette_none:
        raise ValueError("Use either --palette or --palette-none, not both.")
    if args.font and args.fonts_none:
        raise ValueError("Use either --font or --fonts-none, not both.")


def _slide_parts_for_numbers(pkg: OOXMLPackage, slide_numbers: set[int]) -> list[str]:
    parts = slide_parts_in_order(pkg)
    selected = []
    for num in sorted(slide_numbers):
        if num < 1 or num > len(parts):
            raise ValueError("Slide number out of range")
        selected.append(parts[num - 1])
    return selected


def _resolve_image_box(args: argparse.Namespace, cx: int, cy: int) -> tuple[int, int, int, int]:
    x = args.x if args.x is not None else 0
    y = args.y if args.y is not None else 0
    w = args.w if args.w is not None else None
    h = args.h if args.h is not None else None

    if args.units == "emu":
        return int(x), int(y), int(w if w is not None else cx), int(h if h is not None else cy)

    factor = 914400
    x_emu = int(x * factor)
    y_emu = int(y * factor)
    w_emu = int(w * factor) if w is not None else cx
    h_emu = int(h * factor) if h is not None else cy
    return (x_emu, y_emu, w_emu, h_emu)


def _load_template(input_path: str | None) -> PotxTemplate:
    return PotxTemplate.open(input_path) if input_path else PotxTemplate.new()


def _example_palette() -> dict[str, Any]:
    return {
        "dark1": "#FFFFFF",
        "light1": "#0B0B0E",
        "dark2": "#2C2C34",
        "light2": "#E9ECF2",
        "accent1": "#1F6BFF",
        "accent2": "#E0328C",
        "accent3": "#F6A225",
        "accent4": "#6B3AF6",
        "accent5": "#38D3FF",
        "accent6": "#FF4D6D",
        "hlink": "#1F6BFF",
        "folHlink": "#C0186B",
        "majorFont": "Aptos Display",
        "minorFont": "Aptos",
    }


def _normalize_report(result: NormalizeResult) -> dict[str, Any]:
    return {
        "slides_total": result.slides_total,
        "slides_touched": result.slides_touched,
        "replacements": result.replacements,
        "per_slide": result.per_slide,
    }


def _audit_report(report: AuditReport) -> dict[str, Any]:
    return {
        "slides_total": report.slides_total,
        "slides_audited": report.slides_audited,
        "per_slide": report.per_slide,
        "masters": report.masters,
        "layouts": report.layouts,
        "groups": report.groups,
        "theme": report.theme,
        "group_by": report.group_by,
    }


def _print_normalize_result(result: NormalizeResult) -> None:
    print(
        "Replacements: {replacements} across {slides} slide(s)".format(
            replacements=result.replacements, slides=result.slides_touched
        )
    )
    if result.per_slide:
        print("Per slide:")
        for slide, count in sorted(result.per_slide.items()):
            print(f"- {slide}: {count}")


def _print_audit_summary(report: AuditReport, *, details: bool) -> None:
    print("Legend:")
    print("- srgbClr: hard-coded hex color")
    print("- schemeClr: theme-based color slot (accent1, dk1, etc.)")
    print("- sysClr: system color (Office-defined)")
    print("- clrMapOvr: slide overrides the master color mapping")
    print("- custom_bg: slide has a background override")
    print("- images: count of pictures or bitmap fills on the slide")
    print("")

    if report.theme:
        print("Theme:")
        theme = report.theme
        print(f"- part: {theme.get('part', '')}")
        print(f"- theme name: {theme.get('theme_name', '') or '(unset)'}")
        print(f"- color scheme: {theme.get('color_scheme_name', '') or '(unset)'}")
        print(f"- font scheme: {theme.get('font_scheme_name', '') or '(unset)'}")
        print("")

    print(f"Slides audited: {report.slides_audited}/{report.slides_total}")
    if report.masters:
        print("Masters:")
        for part, stats in report.masters.items():
            print(
                f"- {part}: srgb={stats['color_counts']['srgb']}, "
                f"scheme={stats['color_counts']['scheme']}, "
                f"fills={stats['fills']['solid']}/{stats['fills']['grad']}/"
                f"{stats['fills']['blip']}, pics={stats['pictures']}"
            )
    if report.layouts:
        print("Layouts:")
        for part, stats in report.layouts.items():
            print(
                f"- {part}: srgb={stats['color_counts']['srgb']}, "
                f"scheme={stats['color_counts']['scheme']}, "
                f"fills={stats['fills']['solid']}/{stats['fills']['grad']}/"
                f"{stats['fills']['blip']}, pics={stats['pictures']}"
            )

    if report.groups:
        print(f"Groups (group_by={','.join(report.group_by)}):")
        for group in report.groups:
            slides = _format_slide_ranges(group["slides"])
            palette = _format_palette(group["palette"])
            print(f"- slides: {slides}")
            if group.get("layout_part"):
                print(f"  layout: {group['layout_part']}")
            if group.get("master_part"):
                print(f"  master: {group['master_part']}")
            if group.get("background"):
                print(f"  background: {group['background']}")
            if palette:
                print(f"  palette: {palette}")
            print(
                f"  hardcoded_total={group['hardcoded_total']} "
                f"(text={group['text_srgb_total']}, shape={group['shape_srgb_total']})"
            )
            if group["clrMapOvr_slides"]:
                print(f"  clrMapOvr slides: {group['clrMapOvr_slides']}")
            if group["custom_bg_slides"]:
                print(f"  custom_bg slides: {group['custom_bg_slides']}")
            if group["image_slides"]:
                print(f"  image slides: {group['image_slides']}")

        _print_group_recommendations(report.groups)

    if not details:
        return

    for slide in sorted(report.per_slide):
        data = report.per_slide[slide]
        counts = data["color_counts"]
        text_counts = data["text_colors"]
        shape_counts = data["shape_colors"]
        fills = data["fills"]
        bg = data["background"]
        flags = []
        hardcoded = counts["srgb"] + counts["sysclr"]
        hardcoded_details = []
        if text_counts["srgb"]:
            hardcoded_details.append(f"text={text_counts['srgb']}")
        if shape_counts["srgb"]:
            hardcoded_details.append(f"shape={shape_counts['srgb']}")
        if hardcoded:
            if hardcoded_details:
                flags.append(
                    f"hardcoded={hardcoded} ({', '.join(hardcoded_details)})"
                )
            else:
                flags.append(f"hardcoded={hardcoded}")
        if data["has_clrMapOvr"]:
            flags.append("clrMapOvr")
        if data["pictures"] or fills["blip"]:
            flags.append(f"images={data['pictures']}")
        if bg["bg_blip"] or bg["bg_grad"] or bg["bg_solid"] or bg["bg_ref"]:
            flags.append("custom_bg")

        top_colors = _format_top_colors(data["top_srgb"])
        summary = ", ".join(flags) if flags else "no overrides detected"
        slide_part = data.get("slide_part", "")
        layout_part = data.get("layout_part", "")
        master_part = data.get("master_part", "")
        print(f"- slide {slide}: {summary}")
        if slide_part:
            print(f"  part: {slide_part}")
        if layout_part:
            print(f"  layout: {layout_part}")
        if master_part:
            print(f"  master: {master_part}")
        if top_colors:
            print(f"  top colors: {top_colors}")
        top_sizes = _format_top_sizes(data.get("text_styles", {}).get("top_sizes", []))
        if top_sizes:
            print(f"  top sizes: {top_sizes}")


def _format_top_colors(entries: list[dict[str, int]]) -> str:
    return ", ".join(
        [f"#{entry['value']}({entry['count']})" for entry in entries]
    )


def _format_palette(values: list[str]) -> str:
    return ", ".join([f"#{value}" for value in values])


def _format_top_sizes(entries: list[dict[str, float]]) -> str:
    parts = []
    for entry in entries:
        pt = entry.get("pt")
        count = entry.get("count")
        if pt is None or count is None:
            continue
        parts.append(f"{pt:g}pt({count})")
    return ", ".join(parts)


def _format_slide_ranges(slides: list[int]) -> str:
    if not slides:
        return ""
    ranges = []
    start = prev = slides[0]
    for num in slides[1:]:
        if num == prev + 1:
            prev = num
            continue
        ranges.append(_format_range(start, prev))
        start = prev = num
    ranges.append(_format_range(start, prev))
    return ", ".join(ranges)


def _format_range(start: int, end: int) -> str:
    return f"{start}-{end}" if start != end else str(start)


def _print_group_recommendations(groups: list[dict[str, Any]]) -> None:
    by_layout: dict[str, list[dict[str, Any]]] = {}
    for group in groups:
        layout = group.get("layout_part") or "(none)"
        by_layout.setdefault(layout, []).append(group)

    recommendations = []
    for layout, items in by_layout.items():
        if len(items) > 1:
            recommendations.append((layout, items))

    if not recommendations:
        return

    print("Recommendations:")
    for layout, items in recommendations:
        palettes = [group.get("palette", []) for group in items]
        slide_sets = [_format_slide_ranges(group["slides"]) for group in items]
        print(
            f"- layout {layout} has {len(items)} palettes; "
            "consider splitting into separate layouts."
        )
        for palette, slides in zip(palettes, slide_sets):
            palette_text = _format_palette(palette)
            print(f"  slides {slides}: {palette_text or '(no palette detected)'}")


def _parse_group_by(value: str | None) -> list[str]:
    if not value:
        return ["p", "l"]
    tokens = [token.strip() for token in value.split(",") if token.strip()]
    if len(tokens) == 1 and len(tokens[0]) > 1:
        tokens = list(tokens[0])
    valid = {"p", "b", "l"}
    result = []
    for token in tokens:
        if token not in valid:
            raise ValueError("group-by must be a combination of p,b,l")
        if token not in result:
            result.append(token)
    return result

def _set_if_present(func, mapping: dict[str, Any], key: str) -> int:
    if key in mapping and mapping[key] is not None:
        func(mapping[key])
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
