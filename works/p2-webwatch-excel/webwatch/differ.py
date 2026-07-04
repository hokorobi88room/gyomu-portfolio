"""前回スナップショットとの差分検知。"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Diff:
    """1ターゲットの差分。行の同一性は全フィールドのタプルで判定する。"""

    added: list[dict[str, str]] = field(default_factory=list)
    removed: list[dict[str, str]] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        return bool(self.added or self.removed)


def _key(row: dict[str, str]) -> tuple[tuple[str, str], ...]:
    return tuple(sorted(row.items()))


def diff_rows(previous: list[dict[str, str]], current: list[dict[str, str]]) -> Diff:
    prev_keys = {_key(r) for r in previous}
    curr_keys = {_key(r) for r in current}
    return Diff(
        added=[r for r in current if _key(r) not in prev_keys],
        removed=[r for r in previous if _key(r) not in curr_keys],
    )


class SnapshotStore:
    """ターゲットごとの前回結果をJSONで保存・読込する。"""

    def __init__(self, directory: str | Path) -> None:
        self.dir = Path(directory)
        self.dir.mkdir(parents=True, exist_ok=True)

    def _path(self, name: str) -> Path:
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in name)
        return self.dir / f"{safe}.json"

    def load(self, name: str) -> list[dict[str, str]]:
        p = self._path(name)
        if not p.exists():
            return []
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            return data if isinstance(data, list) else []
        except json.JSONDecodeError:
            return []  # 壊れたスナップショットは初回扱い(全件added)

    def save(self, name: str, rows: list[dict[str, str]]) -> None:
        self._path(name).write_text(
            json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8"
        )
