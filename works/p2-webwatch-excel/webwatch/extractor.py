"""HTMLからの項目抽出(CSSセレクタ)。"""
from __future__ import annotations

from bs4 import BeautifulSoup

from .config import Target


def extract(html: str, target: Target) -> list[dict[str, str]]:
    """item_selector で1件分のブロックを列挙し、fields の相対セレクタで値を取る。

    セレクタに合致しないフィールドは空文字(行ごと欠落させない)。
    """
    soup = BeautifulSoup(html, "html.parser")
    rows: list[dict[str, str]] = []
    for block in soup.select(target.item_selector):
        row: dict[str, str] = {}
        for column, selector in target.fields.items():
            el = block.select_one(selector)
            row[column] = el.get_text(strip=True) if el else ""
        if any(v for v in row.values()):  # 全フィールド空のブロックは除外
            rows.append(row)
    return rows
