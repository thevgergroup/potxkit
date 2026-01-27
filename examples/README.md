# Examples

These scripts demonstrate common potxkit workflows. Run them with Poetry so the local package is available on the path.

## Setup

```bash
poetry install
```

## Scripts

- `basic_edit.py`: edit theme colors and fonts in an existing `.potx`.
- `new_template.py`: create a new template from the bundled base.
- `palette_json.py`: apply a palette from `palette.json`.
- `mapping.json`: sample color-to-theme mapping for `normalize`.
 
## CLI usage

The same workflows are available via the CLI:

```bash
poetry run potxkit new output.potx
poetry run potxkit palette-template --pretty > palette.json
poetry run potxkit info path/to/template.potx
poetry run potxkit apply-palette examples/palette.json output.potx --input path/to/template.potx
poetry run potxkit apply-palette examples/palette.json output.potx
poetry run potxkit dump-theme path/to/template.potx --pretty
poetry run potxkit audit path/to/template.pptx --pretty --output audit.json
poetry run potxkit audit path/to/template.pptx --summary
poetry run potxkit audit path/to/template.pptx --summary --details
poetry run potxkit audit path/to/template.pptx --summary --group-by p,b,l
poetry run potxkit normalize examples/mapping.json output.pptx --input path/to/template.pptx --slides 1,3-5
poetry run potxkit set-colors output.potx --input path/to/template.potx --accent1 #1F6BFF --hlink #1F6BFF
poetry run potxkit set-fonts output.potx --input path/to/template.potx --major "Aptos Display" --minor "Aptos"
poetry run potxkit set-theme-names output.potx --input path/to/template.potx --theme "Code Janitor" --colors "Code Janitor Colors" --fonts "Code Janitor Fonts"
poetry run potxkit make-layout --from-slide 7 --name "Layout Bob" --assign-slides 1-7,9-10 output.pptx --input path/to/template.pptx
poetry run potxkit set-layout --layout "Layout Bob" --palette examples/mapping.json output.pptx --input path/to/template.pptx
poetry run potxkit set-slide --slides 1-11 --palette-none --fonts-none output.pptx --input path/to/template.pptx
poetry run potxkit set-text-styles --layout "Layout Bob" --title-size 30 --title-bold --body-size 20 --body-regular output.pptx --input path/to/template.pptx
poetry run potxkit set-text-styles --layout "Layout Bob" --from-slide 7 output.pptx --input path/to/template.pptx
poetry run potxkit set-layout-bg --layout "Layout Bob" --image path/to/hero.png output.pptx --input path/to/template.pptx
poetry run potxkit set-layout-image --layout "Layout Bob" --image path/to/overlay.png --x 1 --y 1 --w 3 --h 2 output.pptx --input path/to/template.pptx
poetry run potxkit dump-tree path/to/template.pptx --grouped --text --summary --output tree_summary.txt
```

### Run

```bash
poetry run python examples/basic_edit.py path/to/input.potx path/to/output.potx
poetry run python examples/new_template.py path/to/output.potx
poetry run python examples/palette_json.py examples/palette.json path/to/output.potx
```
