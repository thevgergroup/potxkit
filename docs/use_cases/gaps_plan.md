# Use Case Gaps & Plan

This document tracks feature gaps against the user-facing use cases in `docs/use_cases/use_cases.md` and the plan to close them.

## Use Case 1: New Template (custom theme colors + fonts + sizes)
**Current coverage**
- Theme colors and font families are supported.
- Theme names (theme/color/font scheme names) are supported.
 - Title/body sizes can be applied to layouts/masters via `set-text-styles` (font family remains separate).

**Gaps**
- No automatic detection of font sizes, weights, or spacing.
- No command to set title/body size/weight at the master/layout level.

**Plan**
1) Add font size/weight extraction for slides, layouts, and masters (audit output). (in progress: slide-level sizes now reported; master/layout pending)
2) Add `set-text-styles` to apply title/body sizes + weights to master/layout placeholders. (done)
3) Provide a `styles.json` schema to apply typography settings in one command. (done: `examples/styles.json`)

## Use Case 2: New Template with a title slide (image + colors + fonts/sizes)
**Current coverage**
- We can create layouts from existing slides and copy shapes/images.
 - Layout background and image layers are supported (`set-layout-bg`, `set-layout-image`).

**Gaps**
- No automatic cropping/positioning rules beyond x/y/w/h (manual placement required).

**Plan**
1) Add `set-layout-bg` to set a background image (blipFill in `p:bgPr`). (done)
2) Add `set-layout-image` to place an image shape at a position/size. (done)
3) Reuse typography controls from Use Case 1.

## Use Case 3: Fix an existing slide deck
**Current coverage**
- Audit detects hard-coded colors, text overrides, images, and background overrides.
- Normalize can replace hard-coded colors with theme slots.
- Layouts can be created from slides, and slides can be assigned to layouts.
 - Audit supports grouping by palette/background/layout with `--group-by p,b,l`.

**Gaps**
- No automatic grouping by palette/background/layout combinations.
- No automatic extraction of font sizes/weights for grouping.
- No auto-generation of multiple layouts based on clusters.

**Plan**
1) Add audit grouping options (`--group-by p,b,l`). (done)
2) Add font-size extraction and include it in group criteria. (in progress)
3) Add `auto-layout` to cluster slides and generate layouts per group. (done: `potxkit auto-layout`)
4) Add `migrate` to apply normalization + assign layouts in one step.

## Grouping Options (requested)
Users can control grouping by palette/background/layout with abbreviations:
- `p` = palette (top hard-coded colors)
- `b` = background flags (bg ref/solid/grad/image)
- `l` = slide layout

Example:
```
potxkit audit deck.pptx --summary --group-by p,b,l
```

## Initial Implementation (begun)
- Added audit grouping selector (`--group-by p,b,l`).
- Added layout creation and slide assignment commands.
- Added palette/font stripping controls for layouts and slides.
