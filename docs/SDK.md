# potxkit SDK

This SDK focuses on editing theme colors and fonts in PowerPoint `.potx` files. It operates directly on Open XML parts and preserves existing layout/masters unless you replace the base template.

## Installation

```bash
poetry install
```

## Core API

### `PotxTemplate`

```python
from potxkit import PotxTemplate
```

- `PotxTemplate.open(path: str, *, fs_kwargs: dict | None = None) -> PotxTemplate`
  - Load an existing `.potx` or `.pptx`.
  - Uses `fsspec` for IO; pass `fs_kwargs` for credentials or storage options.
- `PotxTemplate.new() -> PotxTemplate`
  - Create a template from the bundled base `base.potx`.
- `PotxTemplate.save(path: str, *, fs_kwargs: dict | None = None) -> None`
  - Write the updated package. Ensures the theme relationship and content type override exist.
- `PotxTemplate.validate() -> ValidationReport`
  - Validate relationship targets and content types.
- `PotxTemplate.theme -> Theme`
  - Access theme colors and fonts.

## CLI

The CLI is available after installing dependencies with Poetry.

```bash
poetry run potxkit --help
```

Commands:

```bash
poetry run potxkit new output.potx
poetry run potxkit palette-template --pretty > palette.json
poetry run potxkit palette-template --output palette.json
poetry run potxkit info path/to/template.potx
poetry run potxkit apply-palette examples/palette.json output.potx --input path/to/template.potx
poetry run potxkit apply-palette examples/palette.json output.potx
poetry run potxkit validate path/to/template.potx
poetry run potxkit audit path/to/template.pptx --pretty --output audit.json
poetry run potxkit audit path/to/template.pptx --summary
poetry run potxkit audit path/to/template.pptx --summary --details
poetry run potxkit audit path/to/template.pptx --summary --group-by p,b,l
poetry run potxkit dump-theme path/to/template.potx --pretty
poetry run potxkit normalize examples/mapping.json output.pptx --input path/to/template.pptx --slides 1,3-5
poetry run potxkit set-colors output.potx --input path/to/template.potx --accent1 #1F6BFF --hlink #1F6BFF
poetry run potxkit set-fonts output.potx --input path/to/template.potx --major "Aptos Display" --minor "Aptos"
poetry run potxkit set-master --master 1 --palette-none output.pptx --input path/to/template.pptx
poetry run potxkit set-theme-names output.potx --input path/to/template.potx --theme "Code Janitor" --colors "Code Janitor Colors" --fonts "Code Janitor Fonts"
poetry run potxkit make-layout --from-slide 7 --name "Layout Bob" --assign-slides 1-7,9-10 output.pptx --input path/to/template.pptx
poetry run potxkit set-layout --layout "Layout Bob" --palette examples/mapping.json output.pptx --input path/to/template.pptx
poetry run potxkit set-slide --slides 1-11 --palette-none --fonts-none output.pptx --input path/to/template.pptx
poetry run potxkit set-text-styles --layout "Layout Bob" --title-size 30 --title-bold --body-size 20 --body-regular output.pptx --input path/to/template.pptx
poetry run potxkit set-text-styles --layout "Layout Bob" --styles examples/styles.json output.pptx --input path/to/template.pptx
poetry run potxkit set-layout-bg --layout "Layout Bob" --image path/to/hero.png output.pptx --input path/to/template.pptx
poetry run potxkit set-layout-image --layout "Layout Bob" --image path/to/overlay.png --x 1 --y 1 --w 3 --h 2 output.pptx --input path/to/template.pptx
poetry run potxkit prune-layouts output.pptx --input path/to/template.pptx
poetry run potxkit reindex-layouts output.pptx --input path/to/template.pptx
poetry run potxkit sanitize output.pptx --input path/to/template.pptx
poetry run potxkit dump-tree path/to/template.pptx --layout --master --text --pretty --output tree.json
poetry run potxkit dump-tree path/to/template.pptx --grouped --text --pretty --output tree_grouped.json
poetry run potxkit dump-tree path/to/template.pptx --grouped --text --summary --output tree_summary.txt
poetry run potxkit auto-layout output.pptx --input path/to/template.pptx --group-by p,b --strip-colors --strip-fonts
```

### Normalize mappings

The `normalize` command replaces hard-coded slide colors with theme scheme colors. The mapping JSON should map hex colors to scheme names.
Use `--slides` to target specific slide numbers or ranges (1-based).

### Audit report

The `audit` command inspects slides for hard-coded colors, color overrides, and basic fill usage. The JSON report includes per-slide color counts, top `srgbClr` values, and background/fill flags so you can decide which ranges to normalize.
Use `--summary` for a readable console report that flags hard-coded colors, text overrides, images, and background overrides.

### Layout and slide editing

- `make-layout` creates a new layout by copying the shapes from a slide into a new `slideLayout` part and wiring it into the existing master. Use `--assign-slides` to point slides at the new layout.
- `set-master` updates the master with a palette mapping or clears local overrides (`--palette-none`, `--fonts-none`).
- `set-layout` updates a layout with a palette mapping or clears local overrides (`--palette-none`, `--fonts-none`).
- `set-slide` updates a slide directly (palette, fonts, or layout assignment).
- `set-text-styles` sets title/body sizes and bold styles (use `--from-slide` to auto-detect).
- `set-layout-bg` sets a background image on a layout.
- `set-layout-image` adds an image layer to a layout (x/y/w/h in inches unless `--units emu`).
- `prune-layouts` removes unused layout parts from the master and package.
- `reindex-layouts` renumbers layout files and updates slide/master references.
- `sanitize` adds missing slide defaults (clrMapOvr, lstStyle, and bg noFill) so PowerPoint doesn't repair the file.
- `dump-tree` emits a hierarchical JSON view of slide backgrounds, shapes, fills, and text color nodes. Use `--grouped` to emit `slideMaster` / `slideLayout` / `local` blocks per slide. Use `--summary` for a compact text report including fonts and sizes.
- `auto-layout` groups slides and creates layouts per group (optional stripping of colors/fonts).

Notes:
- `--palette-none` removes hard-coded colors so the theme can flow through.
- `--fonts-none` strips inline text formatting to fall back to master styles.

## Glossary (OOXML color nodes)

- `srgbClr`: a hard-coded hex color value on a slide, layout, or master.
- `schemeClr`: a theme-based color slot (e.g., `accent1`, `dk1`) that follows the theme.
- `sysClr`: a system color defined by Office, often with a `lastClr` fallback.
- `clrMap`: color mapping defined by a master.
- `clrMapOvr`: a slide override of the master color mapping.
- `solidFill` / `gradFill` / `blipFill`: solid, gradient, and image fills.

Example mapping (`examples/mapping.json`):

```json
{
  "#0D0D14": "dark1",
  "#FFFFFF": "light1",
  "#00D9FF": "accent1"
}
```

Scheme names accepted: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`-`accent6`, `hlink`, `folHlink`. Aliases `dark1`, `light1`, `dark2`, `light2` are accepted.

### `Theme`

```python
theme = tpl.theme
```

- `theme.colors: ThemeColors`
- `theme.fonts: ThemeFonts`
- `theme.get_name() / theme.set_name(value)`: theme name shown in PowerPoint.
- `theme.get_color_scheme_name() / theme.set_color_scheme_name(value)`
- `theme.get_font_scheme_name() / theme.set_font_scheme_name(value)`

### `ThemeColors`

All colors are hex in `#RRGGBB` format. Invalid values raise `ValueError`.

- `get_dark1() / set_dark1(hex)`
- `get_light1() / set_light1(hex)`
- `get_dark2() / set_dark2(hex)`
- `get_light2() / set_light2(hex)`
- `get_accent(i: int) / set_accent(i: int, hex)` (1–6)
- `get_hyperlink() / set_hyperlink(hex)`
- `get_followed_hyperlink() / set_followed_hyperlink(hex)`
- `as_dict() -> dict[str, str | None]`

### `ThemeFonts`

- `get_major() -> ThemeFontSpec | None`
- `get_minor() -> ThemeFontSpec | None`
- `set_major(latin: str, east_asian: str | None = None, complex_script: str | None = None)`
- `set_minor(latin: str, east_asian: str | None = None, complex_script: str | None = None)`

### `ThemeFontSpec`

```python
ThemeFontSpec(latin: str, east_asian: str | None, complex_script: str | None)
```

### `ValidationReport`

```python
report = tpl.validate()
if not report.ok:
    print(report.errors)
```

- `errors: list[str]`
- `warnings: list[str]`
- `ok: bool`

## Examples

Edit an existing template:

```python
from potxkit import PotxTemplate

tpl = PotxTemplate.open("template.potx")
tpl.theme.colors.set_accent(1, "#1F6BFF")
tpl.theme.fonts.set_major("Aptos Display")
tpl.save("brand-template.potx")
```

Create a new template from the bundled base:

```python
from potxkit import PotxTemplate

tpl = PotxTemplate.new()
tpl.theme.colors.set_accent(2, "#E0328C")
tpl.save("new-theme.potx")
```

Validate a file:

```python
from potxkit import PotxTemplate

tpl = PotxTemplate.open("template.potx")
report = tpl.validate()
print(report.ok, report.errors, report.warnings)
```

## Notes and limitations

- Theme edits currently target `ppt/theme/theme1.xml` only.
- Slide master/layout styling is not authored here; use a PowerPoint-authored base for complex layouts.
- For remote storage, pass `fs_kwargs` (for example, credentials) into `open()` and `save()`.
