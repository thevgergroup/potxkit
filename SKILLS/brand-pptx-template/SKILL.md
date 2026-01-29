---
name: brand-pptx-template
description: Create branded PowerPoint templates (.pptx/.potx) by extracting brand styles from a website or assets, generating base slides, and applying themes/layouts with potxkit via MCP or CLI.
---

# Brand PPTX Template Skill

## When to use
Use when asked to build or refresh a company PowerPoint template based on a website or brand assets, and deliver both `.pptx` and `.potx` outputs.

## Inputs to collect
- Brand source: website URL or provided brand assets (logo, colors, fonts).
- Output scope: slide types needed (title, section, content, etc.).
- Local working directory path (for potxkit access).
- Delivery format: `.pptx` only or `.pptx` + `.potx`.

## Environment model (generic)
Two environments are common:
- **Agent runtime**: where you can generate files (e.g., pptxgenjs, scripts).
- **User local machine**: where potxkit runs (MCP or CLI) and files live.

If the environments differ, always:
1) Generate base template in the agent runtime.
2) Transfer to local (download/present file).
3) Use potxkit on the local machine to apply theme + layouts.

## Workflow (high level)

### Phase 1: Brand extraction
- Use a browser MCP to extract colors, fonts, and logo.
- Capture UI patterns: light/dark sections, accent usage, card styles.
- If blocked, ask for brand assets or a style guide.

Reference: `references/brand-extraction.js` for a browser snippet.

### Variant chooser (light vs dark)
Decide on a template mood before building slides:
- **Light**: default for business decks; white/very light background with dark text.
- **Dark**: good for technical or cinematic decks; dark background with light text.

If the brand site uses both, pick one for the base template and keep the other for section/cover layouts.

### Phase 2: Map to PowerPoint theme slots
Map to theme keys:
- `dark1`, `light1`, `dark2`, `light2`
- `accent1`–`accent6`, `hlink`, `folHlink`
- `majorFont`, `minorFont`

### Phase 3: Build base slides
Create a clean base deck with the needed slide types.
Recommended defaults:
- Title, section, content, two-column, grid, metrics, quote, closing.
- Keep layout shapes clean and consistent; avoid hard-coded colors where possible.

Reference: `references/pptxgenjs-starter.js` for a starting scaffold.
Optional script: `scripts/scaffold_deck.js` to generate a minimal base deck:

```bash
node scripts/scaffold_deck.js --variant light --company "Acme Co" --out template-base.pptx
```

### Phase 4: Apply theme + layouts with potxkit
Use potxkit via MCP or CLI:
- `set_colors` or `apply-palette`
- `set_fonts`
- `set_theme_names`
- `make_layout` from each base slide
- `validate`

### Phase 5: Deliver
- Provide `.pptx` and `.potx` outputs.
- Confirm template behaves correctly in PowerPoint.

## Guardrails
- Prefer theme-driven colors/fonts instead of per-shape overrides.
- Keep font sizes sane (titles 36–44 pt, body 14–18 pt).
- If outputs are generated in a container, always transfer to local before potxkit.

## References
- Brand extraction script: `references/brand-extraction.js`
- PptxGenJS scaffold: `references/pptxgenjs-starter.js`
- Scaffold generator: `scripts/scaffold_deck.js`
