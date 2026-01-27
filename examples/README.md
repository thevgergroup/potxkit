# Examples

These examples are meant to be copy/paste friendly and explain what each command accomplishes so you can adapt them to real decks.

## Setup

```bash
poetry install
```

## MCP server

Use this if you want an agent to call potxkit tools directly:

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

## Scripted workflows

### Edit an existing theme

```bash
poetry run python examples/basic_edit.py path/to/input.potx path/to/output.potx
```

Edits accent colors and font families in a `.potx`.

### Create a fresh template

```bash
poetry run python examples/new_template.py path/to/output.potx
```

Creates a new theme from the bundled base.

### Apply a palette.json

```bash
poetry run python examples/palette_json.py examples/palette.json path/to/output.potx
```

Applies a palette file (use `palette-template` to generate one).

## CLI workflows (with intent)

### Inspect a deck or template

```bash
poetry run potxkit-cli info path/to/template.potx
poetry run potxkit-cli dump-theme path/to/template.potx --pretty
```

Use these to confirm current theme colors and fonts before changing anything.

### Audit where formatting is coming from

```bash
poetry run potxkit-cli audit path/to/template.pptx --summary --group-by p,b,l
poetry run potxkit-cli dump-tree path/to/template.pptx --grouped --text --summary --output tree_summary.txt
```

`audit` summarizes local overrides; `dump-tree` shows master/layout/slide sources.

### Standardize palette and fonts

```bash
poetry run potxkit-cli set-colors output.potx --input path/to/template.potx --accent1 #1F6BFF
poetry run potxkit-cli set-fonts output.potx --input path/to/template.potx --major "Aptos Display"
```

Use for simple theme updates without touching layouts.

### Create a layout from a slide

```bash
poetry run potxkit-cli make-layout --from-slide 7 --name "Layout Bob" \
  --assign-slides 1-7,9-10 output.pptx --input path/to/template.pptx
```

Copies shapes/images from slide 7 into a new layout and assigns it to a slide range.

### Remove local overrides (let layouts drive style)

```bash
poetry run potxkit-cli set-slide --slides 1-11 --palette-none --fonts-none \
  output.pptx --input path/to/template.pptx
```

Use when the deck has heavy inline formatting and you want inheritance from layouts/master.

### Normalize mixed palettes

```bash
poetry run potxkit-cli normalize examples/mapping.json output.pptx \
  --input path/to/template.pptx --slides 1,3-5
```

Maps arbitrary slide colors to a standard theme palette.
