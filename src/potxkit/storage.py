from __future__ import annotations

from typing import Any

import fsspec


def read_bytes(uri: str, *, fs_kwargs: dict[str, Any] | None = None) -> bytes:
    fs_kwargs = fs_kwargs or {}
    with fsspec.open(uri, "rb", **fs_kwargs) as handle:
        return handle.read()


def write_bytes(
    uri: str, data: bytes, *, fs_kwargs: dict[str, Any] | None = None
) -> None:
    fs_kwargs = fs_kwargs or {}
    with fsspec.open(uri, "wb", **fs_kwargs) as handle:
        handle.write(data)
