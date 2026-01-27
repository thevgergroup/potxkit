#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR=$(cd "$(dirname "$0")/.." && pwd)
git -C "$ROOT_DIR" config core.hooksPath .githooks
echo "Git hooks installed (core.hooksPath=.githooks)"
