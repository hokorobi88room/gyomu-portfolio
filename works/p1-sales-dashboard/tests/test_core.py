"""SalesLens コアロジックのテスト(正常系+異常系+性能)。"""
from __future__ import annotations

import io
import sys
import time
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import core  # noqa: E402


def make_csv(text: str) -> bytes:
    return text.encode("utf-8")


# ---------- 正常系 ----------

def test_normalize_and_clean_happy_path() -> None:
    raw = core.read_table(
        make_csv("日付,商品名,地域,売上(円)\n2026-01-10,プランA,東京,1000\n2026-01-11,プランB,大阪,2000\n"),
        "sales.csv",
    )
    df = core.normalize_columns(raw)
    result = core.clean(df)
    assert result.kept_rows == 2
    assert result.dropped == {}
    assert set(df.columns) >= {"date", "product", "region", "amount"}


def test_column_alias_absorption() -> None:
    # 「売上高」「品目」「取引日」でも読める
    raw = core.read_table(
        make_csv("取引日,品目,売上高\n2026-02-01,保守,3000\n"), "x.csv"
    )
    df = core.normalize_columns(raw)
    assert core.clean(df).kept_rows == 1
    assert (df["region"] == "(未分類)").all()  # 地域列なし → 未分類


def test_cp932_csv() -> None:
    data = "日付,商品,売上\n2026-03-01,テスト商品,5000\n".encode("cp932")
    df = core.normalize_columns(core.read_table(data, "sjis.csv"))
    assert core.clean(df).kept_rows == 1


def test_monthly_summary_yoy() -> None:
    df = pd.DataFrame({
        "date": pd.to_datetime(["2025-01-15", "2025-01-20", "2026-01-10"]),
        "product": ["A", "A", "A"],
        "region": ["東京"] * 3,
        "amount": [100.0, 100.0, 300.0],
    })
    m = core.monthly_summary(df)
    jan26 = m[m["month"] == "2026-01"].iloc[0]
    assert jan26["amount"] == 300.0
    assert jan26["yoy"] == pytest.approx(0.5)  # 200 → 300 は +50%


def test_breakdown_top_n_and_other() -> None:
    df = pd.DataFrame({
        "date": pd.to_datetime(["2026-01-01"] * 12),
        "product": [f"P{i}" for i in range(12)],
        "region": ["東京"] * 12,
        "amount": [float(100 - i) for i in range(12)],
    })
    b = core.breakdown(df, "product", top_n=10)
    assert len(b) == 11
    assert b.iloc[-1]["product"] == "その他"
    assert b.iloc[-1]["amount"] == pytest.approx(89.0 + 90.0)  # P10+P11


def test_detect_anomalies_finds_spike() -> None:
    dates = pd.date_range("2026-01-01", periods=30, freq="D")
    amounts = [1000.0] * 30
    amounts[15] = 100000.0  # 急伸日
    df = pd.DataFrame({
        "date": dates, "product": ["A"] * 30, "region": ["東京"] * 30, "amount": amounts,
    })
    anom = core.detect_anomalies(df, sigma=3.0)
    assert len(anom) == 1
    assert anom.iloc[0]["amount"] == 100000.0


def test_excel_report_roundtrip() -> None:
    df = pd.DataFrame({
        "date": pd.to_datetime(["2026-01-10", "2026-02-10"]),
        "product": ["A", "B"], "region": ["東京", "大阪"], "amount": [1000.0, 2000.0],
    })
    data = core.to_excel_report(df)
    back = pd.read_excel(io.BytesIO(data), sheet_name="月次推移")
    assert list(back["amount"]) == [1000.0, 2000.0]


# ---------- 異常系 ----------

def test_missing_required_column_raises() -> None:
    raw = core.read_table(make_csv("日付,数量\n2026-01-01,3\n"), "bad.csv")
    with pytest.raises(core.DataFormatError, match="必須列"):
        core.normalize_columns(raw)


def test_broken_rows_are_counted_not_silently_dropped() -> None:
    raw = core.read_table(
        make_csv(
            "日付,商品,売上\n"
            "2026-01-10,プランA,1000\n"
            "こわれた日付,プランB,2000\n"
            "2026-01-12,プランC,数字じゃない\n"
            "2026-01-13,,3000\n"
        ),
        "dirty.csv",
    )
    result = core.clean(core.normalize_columns(raw))
    assert result.kept_rows == 1
    assert result.dropped == {"日付が不正": 1, "売上が数値でない": 1, "商品名が空": 1}


def test_few_days_no_anomaly_detection() -> None:
    # 標本7日未満では誤検知防止のため検知しない
    df = pd.DataFrame({
        "date": pd.to_datetime(["2026-01-01", "2026-01-02"]),
        "product": ["A", "A"], "region": ["東京"] * 2, "amount": [1.0, 9999999.0],
    })
    assert core.detect_anomalies(df).empty


# ---------- 性能(10万行) ----------

def test_100k_rows_under_3_seconds() -> None:
    n = 100_000
    df = pd.DataFrame({
        "date": pd.to_datetime("2024-01-01") + pd.to_timedelta(range(n), unit="m"),
        "product": [f"P{i % 20}" for i in range(n)],
        "region": [f"R{i % 5}" for i in range(n)],
        "amount": [float(1000 + i % 977) for i in range(n)],
    })
    t0 = time.perf_counter()
    core.monthly_summary(df)
    core.breakdown(df, "product")
    core.detect_anomalies(df)
    core.kpis(df)
    elapsed = time.perf_counter() - t0
    assert elapsed < 3.0, f"10万行の分析に {elapsed:.2f} 秒かかりました(基準3秒)"
