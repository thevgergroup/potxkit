#!/usr/bin/env python3
"""Build the potxkit MCP bundle (.mcpb)."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = ROOT / "mcpb"
SERVER_SRC = TEMPLATE_DIR / "src" / "server.py"


def _resolve_version(explicit: str | None) -> str:
    if explicit:
        return explicit.lstrip("v")
    env_version = os.getenv("POTXKIT_VERSION")
    if env_version:
        return env_version.lstrip("v")
    ref_name = os.getenv("GITHUB_REF_NAME")
    if ref_name:
        return ref_name.lstrip("v")
    try:
        result = subprocess.run(
            ["git", "describe", "--tags", "--abbrev=0"],
            check=True,
            capture_output=True,
            text=True,
            cwd=ROOT,
        )
        return result.stdout.strip().lstrip("v")
    except subprocess.CalledProcessError as exc:
        raise RuntimeError("Unable to determine version; pass --version.") from exc


def _write_manifest(dest: Path, version: str) -> None:
    manifest = {
        "manifest_version": "0.4",
        "name": "potxkit",
        "display_name": "potxkit",
        "version": version,
        "description": "PowerPoint theme and template fixer for MCP",
        "author": {
            "name": "Patrick O'Leary",
        },
        "server": {
            "type": "uv",
            "entry_point": "src/server.py",
            "mcp_config": {
                "command": "python",
                "args": ["${__dirname}/src/server.py"],
            },
        },
        "compatibility": {
            "platforms": ["darwin", "linux", "win32"],
            "runtimes": {"python": ">=3.11"},
        },
        "keywords": ["powerpoint", "templates", "mcp", "pptx", "potx"],
        "license": "MIT",
    }
    dest.write_text(json.dumps(manifest, indent=2) + "\n")


def _write_pyproject(dest: Path, version: str) -> None:
    content = (
        "[project]\n"
        'name = "potxkit-mcp"\n'
        f'version = "{version}"\n'
        'description = "potxkit MCP server"\n'
        'requires-python = ">=3.11"\n'
        "dependencies = [\n"
        f'    "potxkit=={version}",\n'
        "]\n"
    )
    dest.write_text(content)


def _ensure_mcpb_cli() -> None:
    if shutil.which("mcpb"):
        return
    raise RuntimeError(
        "mcpb CLI not found. Install with: npm install -g @anthropic-ai/mcpb"
    )


def build_bundle(output: Path, version: str) -> None:
    workdir = output.parent / "mcpb"
    if workdir.exists():
        shutil.rmtree(workdir)
    (workdir / "src").mkdir(parents=True, exist_ok=True)

    shutil.copy2(SERVER_SRC, workdir / "src" / "server.py")
    _write_manifest(workdir / "manifest.json", version)
    _write_pyproject(workdir / "pyproject.toml", version)

    _ensure_mcpb_cli()
    output.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(["mcpb", "pack", str(workdir), str(output)], check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build potxkit MCP bundle")
    parser.add_argument("--version", help="Bundle version (e.g., 0.2.1)")
    parser.add_argument(
        "--output",
        default=str(ROOT / "dist" / "potxkit.mcpb"),
        help="Output .mcpb path",
    )
    args = parser.parse_args()

    version = _resolve_version(args.version)
    output = Path(args.output)
    build_bundle(output, version)
    print(f"Built {output}")


if __name__ == "__main__":
    main()
