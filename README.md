# potxkit

potxkit helps you inspect, clean, and standardize PowerPoint themes and slide styling. It works directly with OOXML parts so you can move local slide formatting back into the master/layouts, apply consistent palettes, and generate branded templates without hand-editing every slide.

## What you can do

- Audit decks to see which slides override the master (colors, text, backgrounds, images).
- Normalize or strip local formatting so slides inherit from layouts/master.
- Create a layout from an existing slide (shapes/images included) and apply it across ranges.
- Apply palettes and font families across templates or decks.
- Reindex/prune layouts and validate OOXML structure.

## Choose your interface

### MCP server (recommended for agents)
Runs over stdio and exposes all commands as tools.

```bash
uvx potxkit
```

Example config (`docs/mcp.json`):

```json
{
  "mcpServers": {
    "potxkit": {
      "command": "uvx",
      "args": ["potxkit"]
    }
  }
}
```

### CLI (recommended for humans)

```bash
poetry run potxkit-cli --help
```

### SDK (recommended for developers)

```python
from potxkit import PotxTemplate

tpl = PotxTemplate.open("template.potx")
tpl.theme.colors.set_accent(1, "#1F6BFF")
tpl.theme.fonts.set_major("Aptos Display")
tpl.save("brand-template.potx")
```

## Common workflows

Audit a deck to find local formatting:

```bash
poetry run potxkit-cli audit path/to/template.pptx --summary --group-by p,b,l
```

Create a new layout from slide 7 and apply it to a range:

```bash
poetry run potxkit-cli make-layout --from-slide 7 --name "Layout Bob" \
  --assign-slides 1-7,9-10 output.pptx --input path/to/template.pptx
```

Strip inline colors/fonts so slides inherit from layouts/master:

```bash
poetry run potxkit-cli set-slide --slides 1-11 --palette-none --fonts-none \
  output.pptx --input path/to/template.pptx
```

Normalize a deck to a palette mapping:

```bash
poetry run potxkit-cli normalize examples/mapping.json output.pptx \
  --input path/to/template.pptx --slides 1,3-5
```

Update theme colors, fonts, and friendly names:

```bash
poetry run potxkit-cli set-colors output.potx --input path/to/template.potx \
  --accent1 #1F6BFF --hlink #1F6BFF
poetry run potxkit-cli set-fonts output.potx --input path/to/template.potx \
  --major "Aptos Display" --minor "Aptos"
poetry run potxkit-cli set-theme-names output.potx --input path/to/template.potx \
  --theme "Code Janitor" --colors "Code Janitor Colors" --fonts "Code Janitor Fonts"
```

## Examples

See `examples/README.md` for step-by-step walkthroughs and the reason each command exists.

## SDK documentation

Full API reference in `docs/SDK.md`.

## Project layout

- `src/potxkit/`: library + MCP/CLI implementation
- `examples/`: runnable scripts and sample inputs
- `docs/`: reference docs and MCP config
- `tests/`: pytest suite

## Notes and limitations

- potxkit edits theme data and slide/layout formatting. It does not render slides.
- For complex branded layouts, start from a PowerPoint-authored `.potx` and use potxkit to standardize themes and remove local overrides.
