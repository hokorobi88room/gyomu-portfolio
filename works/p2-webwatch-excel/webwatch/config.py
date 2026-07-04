"""設定ファイル(YAML)の読み込みと検証。"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml


class ConfigError(ValueError):
    """設定ファイルの形式・内容の問題。"""


@dataclass(frozen=True)
class Target:
    """観測対象1件。"""

    name: str
    url: str
    item_selector: str                 # 1件分のブロックを指すCSSセレクタ
    fields: dict[str, str]             # 列名 → ブロック内の相対CSSセレクタ


@dataclass(frozen=True)
class Config:
    targets: list[Target]
    interval_seconds: float = 5.0      # 同一ドメインへの最小アクセス間隔
    timeout_seconds: float = 15.0
    user_agent: str = "WebWatchBot/1.0 (monitoring; contact via site form)"
    output_dir: Path = field(default_factory=lambda: Path("output"))


def load_config(path: str | Path) -> Config:
    """YAMLを読み込み、必須項目と値域を検証して Config を返す。"""
    p = Path(path)
    if not p.exists():
        raise ConfigError(f"設定ファイルがありません: {p}")

    try:
        raw = yaml.safe_load(p.read_text(encoding="utf-8"))
    except yaml.YAMLError as e:
        raise ConfigError(f"YAMLの構文エラー: {e}") from e
    if not isinstance(raw, dict):
        raise ConfigError("設定ファイルの最上位はマッピングである必要があります")

    targets_raw = raw.get("targets")
    if not isinstance(targets_raw, list) or not targets_raw:
        raise ConfigError("targets に観測対象を1件以上定義してください")

    targets: list[Target] = []
    for i, t in enumerate(targets_raw):
        for key in ("name", "url", "item_selector", "fields"):
            if key not in t:
                raise ConfigError(f"targets[{i}] に {key} がありません")
        if not str(t["url"]).startswith(("http://", "https://")):
            raise ConfigError(f"targets[{i}].url はhttp(s)で始まる必要があります: {t['url']}")
        if not isinstance(t["fields"], dict) or not t["fields"]:
            raise ConfigError(f"targets[{i}].fields は列名→セレクタのマッピングです")
        targets.append(
            Target(
                name=str(t["name"]),
                url=str(t["url"]),
                item_selector=str(t["item_selector"]),
                fields={str(k): str(v) for k, v in t["fields"].items()},
            )
        )

    interval = float(raw.get("interval_seconds", 5.0))
    if interval < 1.0:
        raise ConfigError("interval_seconds は1秒以上にしてください(相手サイトへの配慮)")

    return Config(
        targets=targets,
        interval_seconds=interval,
        timeout_seconds=float(raw.get("timeout_seconds", 15.0)),
        user_agent=str(raw.get("user_agent", Config.user_agent)),
        output_dir=Path(raw.get("output_dir", "output")),
    )
