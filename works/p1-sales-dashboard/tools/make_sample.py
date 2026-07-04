"""サンプル売上データ(架空・3年分)の生成。"""
from __future__ import annotations

import csv
import math
import random
from datetime import date, timedelta
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "sample-data" / "sales_sample.csv"

PRODUCTS = [
    ("スタンダードプラン", 9800, 6),
    ("プレミアムプラン", 29800, 3),
    ("ライトプラン", 4980, 8),
    ("初期構築パック", 50000, 1),
    ("保守サポート", 15000, 4),
    ("オプション追加", 3000, 5),
]
REGIONS = ["東京", "大阪", "名古屋", "福岡", "札幌"]


def main() -> None:
    rng = random.Random(7)
    rows: list[list[str]] = []
    start = date(2023, 7, 1)
    end = date(2026, 6, 30)
    d = start
    while d <= end:
        # 季節性(3月・9月が繁忙)+ゆるやかな成長トレンド
        season = 1.0 + 0.35 * math.sin((d.month - 6) / 12 * 2 * math.pi)
        growth = 1.0 + (d - start).days / (end - start).days * 0.6
        n = max(1, int(rng.gauss(6 * season * growth, 2)))
        # 異常日を意図的に混ぜる(デモで検知させるため)
        if d in (date(2025, 3, 14), date(2026, 3, 11)):
            n *= 6  # キャンペーン急伸
        if d == date(2025, 11, 5):
            n = 0  # システム障害で売上ゼロ
        for _ in range(n):
            name, price, w = rng.choices(PRODUCTS, weights=[p[2] for p in PRODUCTS])[0]
            qty = rng.choice([1, 1, 1, 2])
            rows.append([
                d.isoformat(),
                name,
                rng.choice(REGIONS),
                str(price * qty),
            ])
        d += timedelta(days=1)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["日付", "商品名", "地域", "売上(円)"])  # わざと別名にして吸収デモ
        w.writerows(rows)
    print(f"generated {len(rows)} rows -> {OUT}")


if __name__ == "__main__":
    main()
