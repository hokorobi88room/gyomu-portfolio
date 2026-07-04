"""SalesLens コアロジック(UIから独立した純関数群)。

Streamlit(app.py)から呼ばれるが、pandasのみで完結するため
単体テスト(tests/)で品質を担保できる。
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field

import pandas as pd

# 列名のゆらぎ吸収マップ(現場のExcel/CSVは列名が揃わない前提で設計)
COLUMN_ALIASES: dict[str, list[str]] = {
    "date": ["日付", "売上日", "取引日", "date", "Date"],
    "product": ["商品", "品目", "商品名", "サービス", "product"],
    "region": ["地域", "支店", "エリア", "店舗", "region"],
    "amount": ["売上", "金額", "売上高", "売上金額", "amount", "売上(円)"],
}

REQUIRED = ("date", "product", "amount")


@dataclass
class CleanResult:
    """クレンジング結果。除外行も黙って捨てず件数と理由を保持する。"""

    df: pd.DataFrame
    total_rows: int
    dropped: dict[str, int] = field(default_factory=dict)

    @property
    def kept_rows(self) -> int:
        return len(self.df)


class DataFormatError(ValueError):
    """必須列が特定できない等、入力データ形式の問題。"""


def read_table(data: bytes, filename: str) -> pd.DataFrame:
    """CSV(UTF-8/CP932自動判定)またはExcelを読み込む。"""
    if filename.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(data))
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return pd.read_csv(io.BytesIO(data), encoding=enc)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    raise DataFormatError("CSVの文字コードを判定できません(UTF-8 か Shift_JIS で保存してください)")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """列名のゆらぎを標準名(date/product/region/amount)へ寄せる。

    必須列が見つからない場合は DataFormatError(候補つきメッセージ)。
    """
    rename: dict[str, str] = {}
    for std, aliases in COLUMN_ALIASES.items():
        for col in df.columns:
            base = str(col).strip()
            # "売上(円)" → "売上" のような注記括弧を除去して比較
            for br in ("(", "("):
                if br in base:
                    base = base.split(br)[0]
            if base in aliases and std not in rename.values():
                rename[col] = std
                break

    df = df.rename(columns=rename)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        jp = {"date": "日付", "product": "商品", "amount": "売上"}
        raise DataFormatError(
            "必須列が見つかりません: "
            + ", ".join(jp[m] for m in missing)
            + f"(認識できる列名の例: {', '.join(COLUMN_ALIASES[missing[0]][:3])})"
        )
    if "region" not in df.columns:
        df["region"] = "(未分類)"
    return df


def clean(df: pd.DataFrame) -> CleanResult:
    """型変換と不正行の除外。除外は理由別に計数して返す(黙って捨てない)。"""
    total = len(df)
    dropped: dict[str, int] = {}

    df = df.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    bad_date = int(df["date"].isna().sum())
    if bad_date:
        dropped["日付が不正"] = bad_date
    df = df.dropna(subset=["date"])

    df["amount"] = pd.to_numeric(
        df["amount"].astype(str).str.replace(",", "").str.replace("円", ""),
        errors="coerce",
    )
    bad_amount = int(df["amount"].isna().sum())
    if bad_amount:
        dropped["売上が数値でない"] = bad_amount
    df = df.dropna(subset=["amount"])

    blank_mask = df["product"].isna() | (df["product"].astype(str).str.strip() == "")
    blank_product = int(blank_mask.sum())
    if blank_product:
        dropped["商品名が空"] = blank_product
    df = df[~blank_mask]

    df["product"] = df["product"].astype(str).str.strip()
    df["region"] = df["region"].astype(str).str.strip().replace("", "(未分類)")

    return CleanResult(df=df.reset_index(drop=True), total_rows=total, dropped=dropped)


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    """月次売上と前年同月比。列: month(str), amount, yoy(前年比。前年なしは NaN)。"""
    m = (
        df.assign(month=df["date"].dt.to_period("M"))
        .groupby("month", as_index=False)["amount"]
        .sum()
        .sort_values("month")
    )
    prev = m.set_index("month")["amount"]
    m["yoy"] = [
        (row.amount / prev[row.month - 12] - 1.0) if (row.month - 12) in prev.index and prev[row.month - 12] != 0 else float("nan")
        for row in m.itertuples()
    ]
    m["month"] = m["month"].astype(str)
    return m


def breakdown(df: pd.DataFrame, by: str, top_n: int = 10) -> pd.DataFrame:
    """商品別・地域別などの構成(降順・上位N+「その他」)。"""
    g = df.groupby(by, as_index=False)["amount"].sum().sort_values("amount", ascending=False)
    if len(g) > top_n:
        head = g.head(top_n)
        other = pd.DataFrame([{by: "その他", "amount": g["amount"].iloc[top_n:].sum()}])
        g = pd.concat([head, other], ignore_index=True)
    return g


def detect_anomalies(df: pd.DataFrame, sigma: float = 3.0) -> pd.DataFrame:
    """日次売上の異常値検知(平均±σ×標準偏差の外側)。

    返り値: date, amount, z(標準化スコア)の DataFrame(異常日のみ)。
    """
    daily = df.groupby(df["date"].dt.date)["amount"].sum()
    if len(daily) < 7:  # 標本が少なすぎる場合は検知しない(誤検知防止)
        return pd.DataFrame(columns=["date", "amount", "z"])
    mean, std = daily.mean(), daily.std(ddof=0)
    if std == 0:
        return pd.DataFrame(columns=["date", "amount", "z"])
    z = (daily - mean) / std
    out = pd.DataFrame({"date": daily.index, "amount": daily.values, "z": z.values})
    return out[out["z"].abs() >= sigma].reset_index(drop=True)


def kpis(df: pd.DataFrame) -> dict[str, float]:
    """ヘッダーに出す主要指標。"""
    monthly = df.groupby(df["date"].dt.to_period("M"))["amount"].sum()
    return {
        "total": float(df["amount"].sum()),
        "avg_month": float(monthly.mean()) if len(monthly) else 0.0,
        "best_month": float(monthly.max()) if len(monthly) else 0.0,
        "n_products": int(df["product"].nunique()),
    }


def to_excel_report(df: pd.DataFrame) -> bytes:
    """月次・商品別・地域別を1ブックにまとめたExcelレポートを返す。"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        monthly_summary(df).to_excel(writer, sheet_name="月次推移", index=False)
        breakdown(df, "product").to_excel(writer, sheet_name="商品別", index=False)
        breakdown(df, "region").to_excel(writer, sheet_name="地域別", index=False)
        df.to_excel(writer, sheet_name="クレンジング済データ", index=False)
    return buf.getvalue()
