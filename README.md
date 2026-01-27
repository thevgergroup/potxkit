# potxkit

potxkit is a Python library for reading and editing PowerPoint `.potx` theme data (colors and fonts) with Open XML parts. It includes a bundled base template and helpers to generate brand-ready `.potx` files.

## Features

- Read and write theme colors (accent slots, hyperlink colors, light/dark variants).
- Read and write theme fonts (major/minor Latin).
- Validate OOXML relationships and content type overrides.
- Create new templates from a bundled base.

## Requirements

- Python 3.11+
- Poetry 2.x

## Installation

```bash
poetry install
```

## MCP server (FastMCP)

The MCP server is the default entry point and runs over stdio. After installing with `uvx`:

```bash
uvx potxkit
```

If you want the CLI instead:

```bash
poetry run potxkit-cli --help
```

## Quick start

Edit an existing template:

```python
from potxkit import PotxTemplate

tpl = PotxTemplate.open("template.potx")
tpl.theme.colors.set_accent(1, "#1F6BFF")
tpl.theme.colors.set_hyperlink("#1F6BFF")
tpl.theme.fonts.set_major("Aptos Display")
tpl.theme.fonts.set_minor("Aptos")
tpl.save("brand-template.potx")
```

Create a new template from the bundled base:

```python
from potxkit import PotxTemplate

tpl = PotxTemplate.new()
tpl.theme.colors.set_accent(2, "#E0328C")
tpl.save("new-theme.potx")
```

## Examples

See `examples/README.md` for runnable scripts:

```bash
poetry run python examples/basic_edit.py path/to/input.potx path/to/output.potx
poetry run python examples/new_template.py path/to/output.potx
poetry run python examples/palette_json.py examples/palette.json path/to/output.potx
```

The normalize command uses a color-to-theme mapping; a sample is in `examples/mapping.json`.

## CLI

After `poetry install`, the CLI is available as:

```bash
poetry run potxkit-cli --help
```

Common commands:

```bash
poetry run potxkit-cli new output.potx
poetry run potxkit-cli palette-template --pretty > palette.json
poetry run potxkit-cli palette-template --output palette.json
poetry run potxkit-cli info path/to/template.potx
poetry run potxkit-cli apply-palette examples/palette.json output.potx --input path/to/template.potx
poetry run potxkit-cli apply-palette examples/palette.json output.potx
poetry run potxkit-cli validate path/to/template.potx
poetry run potxkit-cli audit path/to/template.pptx --pretty --output audit.json
poetry run potxkit-cli audit path/to/template.pptx --summary
poetry run potxkit-cli audit path/to/template.pptx --summary --details
poetry run potxkit-cli audit path/to/template.pptx --summary --group-by p,b,l
poetry run potxkit-cli dump-theme path/to/template.potx --pretty
poetry run potxkit-cli normalize examples/mapping.json output.pptx --input path/to/template.pptx --slides 1,3-5
poetry run potxkit-cli set-colors output.potx --input path/to/template.potx --accent1 #1F6BFF --hlink #1F6BFF
poetry run potxkit-cli set-fonts output.potx --input path/to/template.potx --major "Aptos Display" --minor "Aptos"
poetry run potxkit-cli set-master --master 1 --palette-none output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-theme-names output.potx --input path/to/template.potx --theme "Code Janitor" --colors "Code Janitor Colors" --fonts "Code Janitor Fonts"
poetry run potxkit-cli make-layout --from-slide 7 --name "Layout Bob" --assign-slides 1-7,9-10 output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-layout --layout "Layout Bob" --palette examples/mapping.json output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-slide --slides 1-11 --palette-none --fonts-none output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-text-styles --layout "Layout Bob" --title-size 30 --title-bold --body-size 20 --body-regular output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-text-styles --layout "Layout Bob" --from-slide 7 output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-text-styles --layout "Layout Bob" --styles examples/styles.json output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-layout-bg --layout "Layout Bob" --image path/to/hero.png output.pptx --input path/to/template.pptx
poetry run potxkit-cli set-layout-image --layout "Layout Bob" --image path/to/overlay.png --x 1 --y 1 --w 3 --h 2 output.pptx --input path/to/template.pptx
poetry run potxkit-cli prune-layouts output.pptx --input path/to/template.pptx
poetry run potxkit-cli reindex-layouts output.pptx --input path/to/template.pptx
poetry run potxkit-cli sanitize output.pptx --input path/to/template.pptx
poetry run potxkit-cli dump-tree path/to/template.pptx --layout --master --text --pretty --output tree.json
poetry run potxkit-cli dump-tree path/to/template.pptx --grouped --text --pretty --output tree_grouped.json
poetry run potxkit-cli dump-tree path/to/template.pptx --grouped --text --summary --output tree_summary.txt
poetry run potxkit-cli auto-layout output.pptx --input path/to/template.pptx --group-by p,b --strip-colors --strip-fonts
```

## SDK documentation

See `docs/SDK.md` for the full API reference and usage details.

## Git hooks

Prevent `.pptx`/`.potx` files from being committed (confidential assets):

```bash
./scripts/install-hooks.sh
```

## Terminology

If you see `srgbClr`, `schemeClr`, or `sysClr` in audit outputs, these are OOXML color nodes. See the glossary in `docs/SDK.md` for quick definitions.

Tip: Use `potxkit audit --summary` for grouped output, and `--details` when you need the per-slide breakdown.

## API overview

- `PotxTemplate.open(path)`: load an existing `.potx` or `.pptx`.
- `PotxTemplate.new()`: create a new template from the bundled base.
- `PotxTemplate.save(path)`: write the updated template.
- `PotxTemplate.validate()`: return a `ValidationReport` with errors/warnings.
- `tpl.theme.colors`: access to theme color slots.
- `tpl.theme.fonts`: access to theme font specs.

## Project layout

- `src/potxkit/`: library source code.
- `src/potxkit/data/base.potx`: bundled base template.
- `scripts/`: build utilities.
- `examples/`: runnable examples.
- `tests/`: pytest suite.

## Development

- Run tests: `poetry run pytest`
- Rebuild base template: `poetry run python scripts/generate_base_template.py`
- Generate sample theme: `poetry run python scripts/create_default_theme.py`

## Notes and limitations

- Only theme colors and fonts are edited today. Slide master shapes, backgrounds, and layouts are not authored beyond a minimal master background in the bundled base.
- For complex branded layouts, start from a PowerPoint-authored `.potx` and use potxkit to update the theme layer.
