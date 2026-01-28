![potxkit logo](https://raw.githubusercontent.com/thevgergroup/potxkit/main/assets/logo.png)

# potxkit

Make PowerPoint templates consistent without manual, slide-by-slide cleanup.

[![Install Claude Desktop](https://img.shields.io/badge/Claude%20Desktop-Install-blue)](https://github.com/thevgergroup/potxkit/releases/latest/download/potxkit.mcpb)
[![Add to Cursor](https://img.shields.io/badge/Cursor-Add%20to%20Cursor-black)](cursor://anysphere.cursor-deeplink/mcp/install?name=potxkit&config=eyJwb3R4a2l0Ijp7ImNvbW1hbmQiOiJ1dngiLCJhcmdzIjpbInBvdHhraXQiXX19)

## Links

- GitHub: https://github.com/thevgergroup/potxkit
- PyPI: https://pypi.org/project/potxkit/

## Why this exists

PowerPoint templates have been hard to understand and fix for decades. Most decks slowly drift as people paste content and override styles. Many PowerPoint libraries make this worse by embedding colors and fonts directly on each slide, which bypasses the slide master and makes the template effectively useless. potxkit fixes the root problem by moving formatting back into the theme, master, and layouts so the template stays in control.

## PowerPoint styling hierarchy

PowerPoint styling is layered:

1) **Theme**: global colors + fonts for the file.
2) **Slide master**: the base look for all slides.
3) **Layouts**: variations like “Title Slide,” “Section Header,” etc.
4) **Local overrides**: formatting applied directly on a slide or shape.

When local overrides are everywhere, layouts and the master stop controlling the look. potxkit shows you where formatting is coming from and lets you push it back into the master/layouts so the deck behaves like a real template again.

## Why potxkit + AI agents

Using a CLI to fix templates is powerful but not friendly. Running potxkit as an MCP server lets an AI agent apply the right sequence of audits and fixes for you—cleaning overrides, applying palettes, and standardizing layouts in minutes instead of hours.

- Audit where colors/fonts/backgrounds/images are coming from.
- Strip local overrides so layouts and masters drive the look.
- Apply a consistent palette mapping across slides.
- Set theme fonts, sizes, and layout images programmatically.
- Let agents orchestrate the workflow with plain-language instructions.

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

## MCP client setup

potxkit runs as a local MCP server. Most clients accept this config:

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

### One-click installs

- **Claude Desktop**: download `potxkit.mcpb` from GitHub releases and install via Settings -> Extensions -> Advanced Settings -> Install Extension.
  - Suggested release asset: https://github.com/thevgergroup/potxkit/releases/latest/download/potxkit.mcpb
- **Cursor**: Add to Cursor link:
  - cursor://anysphere.cursor-deeplink/mcp/install?name=potxkit&config=eyJwb3R4a2l0Ijp7ImNvbW1hbmQiOiJ1dngiLCJhcmdzIjpbInBvdHhraXQiXX19

### CLI or config installs

**Claude Code**

```bash
claude mcp add --transport stdio potxkit -- uvx potxkit
```

If you already installed the Claude Desktop extension, you can import it:

```bash
claude mcp add-from-claude-desktop
```

**Codex (OpenAI)**

```bash
codex mcp add potxkit -- uvx potxkit
```

Or add to `~/.codex/config.toml` (or project `.codex/config.toml`):

```toml
[mcp_servers.potxkit]
command = "uvx"
args = ["potxkit"]
```

**Gemini CLI**

Add to your project `.gemini/settings.json` under `mcpServers`:

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

**Roo Code**

Add to `.roo/mcp.json` (project) or `mcp_settings.json` (global):

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

### Maintainer release checklist (one-click installs)

- Tag a release (`vX.Y.Z`). GitHub Actions builds and uploads `potxkit.mcpb` using `scripts/build_mcpb.py`.
- For local builds: `python scripts/build_mcpb.py --version X.Y.Z`.
- If the MCP command or args change, update the Cursor deep link config (base64 of `{\"potxkit\":{\"command\":\"uvx\",\"args\":[\"potxkit\"]}}`).
- Click both install links to verify they still open correctly.

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

## Images, palettes, fonts, and sizes

Add a background image to a layout:

```bash
poetry run potxkit-cli set-layout-bg --layout "Layout Bob" \
  --image path/to/hero.png output.pptx --input path/to/template.pptx
```

Add an image layer (x/y/w/h in inches unless `--units emu`):

```bash
poetry run potxkit-cli set-layout-image --layout "Layout Bob" \
  --image path/to/overlay.png --x 1 --y 1 --w 3 --h 2 \
  output.pptx --input path/to/template.pptx
```

Apply a palette mapping:

```bash
poetry run potxkit-cli normalize examples/mapping.json output.pptx \
  --input path/to/template.pptx --slides 1,3-5
```

Full palette file example (for `apply-palette`):

```json
{
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
  "minorFont": "Aptos"
}
```

What the keys mean (simple):

- `dark1`, `light1`, `dark2`, `light2`: base theme colors PowerPoint uses for text/backgrounds.
- `accent1`–`accent6`: the six accent colors used for charts, shapes, and theme color picks.
- `hlink`, `folHlink`: hyperlink and followed‑link colors.
- `majorFont`, `minorFont`: theme font families (major = headings, minor = body).

```bash
poetry run potxkit-cli apply-palette examples/palette.json output.potx \
  --input path/to/template.potx
```

Set fonts:

```bash
poetry run potxkit-cli set-fonts output.potx --input path/to/template.potx \
  --major "Aptos Display" --minor "Aptos"
```

Set text sizes (points) and bold/regular for a layout:

```bash
poetry run potxkit-cli set-text-styles --layout "Layout Bob" \
  --title-size 30 --title-bold --body-size 20 --body-regular \
  output.pptx --input path/to/template.pptx
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

## License

MIT License. See `LICENSE`.
