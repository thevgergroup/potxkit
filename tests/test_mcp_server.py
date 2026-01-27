from __future__ import annotations

import json
from contextlib import asynccontextmanager
from pathlib import Path

import conftest
import pytest
from fastmcp import Client

from potxkit.mcp_server import mcp


def _tool_names(tools) -> set[str]:
    names: set[str] = set()
    for tool in tools:
        if isinstance(tool, dict):
            name = tool.get("name")
        else:
            name = getattr(tool, "name", None)
        if name:
            names.add(name)
    return names


@pytest.fixture
async def mcp_client(monkeypatch):
    @asynccontextmanager
    async def _no_docket():
        yield

    monkeypatch.setattr(mcp, "_docket_lifespan", _no_docket)
    async with Client(transport=mcp) as client:
        yield client


async def test_mcp_list_tools_includes_core(mcp_client) -> None:
    tools = await mcp_client.list_tools()
    names = _tool_names(tools)
    for expected in {"info", "validate", "dump_theme", "audit", "dump_tree"}:
        assert expected in names


async def test_mcp_dump_theme_and_validate(tmp_path: Path, mcp_client) -> None:
    src = tmp_path / "sample.potx"
    src.write_bytes(conftest.build_minimal_potx())

    result = await mcp_client.call_tool(
        name="dump_theme", arguments={"path": str(src), "pretty": True}
    )
    payload = json.loads(result.data)
    assert payload.get("dk1") is not None

    result = await mcp_client.call_tool(name="validate", arguments={"path": str(src)})
    assert result.data["ok"] is True


async def test_mcp_set_theme_names(tmp_path: Path, mcp_client) -> None:
    src = tmp_path / "sample.potx"
    out = tmp_path / "named.potx"
    src.write_bytes(conftest.build_minimal_potx())

    result = await mcp_client.call_tool(
        name="set_theme_names",
        arguments={
            "input_path": str(src),
            "output": str(out),
            "theme": "Test Theme",
            "colors": "Test Colors",
            "fonts": "Test Fonts",
        },
    )
    assert out.exists()
    assert result.data == str(out)
