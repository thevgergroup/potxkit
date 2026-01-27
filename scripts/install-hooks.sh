#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR=$(cd "$(dirname "$0")/.." && pwd)
git -C "$ROOT_DIR" config core.hooksPath .githooks
if command -v pre-commit >/dev/null 2>&1; then
  # pre-commit refuses to install when core.hooksPath is set; our .githooks/pre-commit
  # script runs pre-commit directly. We still warm the environments here.
  (cd "$ROOT_DIR" && pre-commit run --hook-stage pre-commit --all-files)
else
  echo "pre-commit not found; install it with 'poetry run pip install pre-commit' or 'pip install pre-commit'."
fi
echo "Git hooks installed (core.hooksPath=.githooks)"
