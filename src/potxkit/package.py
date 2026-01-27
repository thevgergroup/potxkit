from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from typing import Iterable


@dataclass(frozen=True)
class PackagePart:
    name: str
    data: bytes


class OOXMLPackage:
    def __init__(self, data: bytes) -> None:
        self._parts: dict[str, bytes] = {}
        self._order: list[str] = []
        self._load(data)

    def _load(self, data: bytes) -> None:
        with zipfile.ZipFile(io.BytesIO(data), "r") as zin:
            for info in zin.infolist():
                self._order.append(info.filename)
                self._parts[info.filename] = zin.read(info.filename)

    def list_parts(self) -> list[str]:
        return list(self._parts.keys())

    def has_part(self, name: str) -> bool:
        return self._normalize_name(name) in self._parts

    def read_part(self, name: str) -> bytes:
        key = self._normalize_name(name)
        if key not in self._parts:
            raise KeyError(f"Part not found: {name}")
        return self._parts[key]

    def write_part(self, name: str, data: bytes) -> None:
        key = self._normalize_name(name)
        if key not in self._parts:
            self._order.append(key)
        self._parts[key] = data

    def delete_part(self, name: str) -> None:
        key = self._normalize_name(name)
        if key in self._parts:
            del self._parts[key]
        self._order = [entry for entry in self._order if entry != key]

    def iter_parts(self) -> Iterable[PackagePart]:
        for name in self._order:
            if name in self._parts:
                yield PackagePart(name=name, data=self._parts[name])

    def save_bytes(self) -> bytes:
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for part in self.iter_parts():
                zout.writestr(part.name, part.data)
                written.add(part.name)
            for name, data in self._parts.items():
                if name not in written:
                    zout.writestr(name, data)
        return out.getvalue()

    @staticmethod
    def _normalize_name(name: str) -> str:
        if name.startswith("/"):
            return name[1:]
        return name
