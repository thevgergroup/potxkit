# Repository Guidelines

## Project Purpose
This repository contains a Python library for editing PowerPoint `.potx` themes (colors and fonts), plus the product requirements that guide its scope.

## Project Structure & Module Organization
- `src/potxkit/`: library source (package, OOXML handling, theme editing).
- `tests/`: pytest suite and helpers.
- `scripts/`: utilities for generating bundled and example templates.
- `templates/`: generated `.potx` outputs for reference.
- `examples/`: runnable scripts showing common workflows.
- `docs/PRD/requirements.md`: primary product requirements and technical notes.
- `.vscode/settings.json`: editor preferences only.

## Build, Test, and Development Commands
- `poetry install`: install dependencies into the local virtual environment.
- `poetry run pytest`: run the test suite.
- `poetry run python scripts/generate_base_template.py`: rebuild the bundled base template.
- `poetry run python scripts/create_default_theme.py`: generate the sample theme output.

## Coding Style & Naming Conventions
- Use 4-space indentation and `snake_case` for Python modules and functions.
- Prefer explicit, short helper functions over deep nesting.
- Keep Markdown sections short and task-focused.

For documentation edits:
- Use Markdown headings and keep sections short and task-focused.
- Use concrete examples and prefer relative paths like `docs/PRD/requirements.md`.

## Testing Guidelines
- Tests use `pytest` and live under `tests/` with `test_*.py` naming.
- Run locally with `poetry run pytest`.
Coverage targets are not defined yet; add them here if requirements change.

## Commit & Pull Request Guidelines
This checkout does not include a Git history, so no commit message conventions are observable. Until a convention is established, use clear, imperative commit messages (e.g., `Add theme XML parsing notes`).

For pull requests:
- Include a concise summary of changes and the motivation.
- Link related issues or design docs in `docs/`.
- Add screenshots or sample outputs when documentation or formatting changes are visual.

## Security & Configuration Tips
Do not commit sample `.potx` files that include sensitive branding or customer data. Use sanitized or synthetic examples instead.
