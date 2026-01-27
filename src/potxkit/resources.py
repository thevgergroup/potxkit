from __future__ import annotations

from importlib import resources


def load_base_template() -> bytes:
    return resources.files("potxkit.data").joinpath("base.potx").read_bytes()
