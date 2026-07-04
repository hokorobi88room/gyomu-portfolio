"""SalesLens — 売上データ分析ダッシュボード(Streamlit)。

CSV/Excelをアップロードすると、クレンジング → KPI → 月次推移(前年比) →
商品別/地域別 → 異常日検知 → Excelレポート出力まで自動で行う。

起動: streamlit run app.py
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import core

st.set_page_config(page_title="SalesLens — 売上分析ダッシュボード", page_icon="📊", layout="wide")

st.title("📊 SalesLens — 売上分析ダッシュボード")
st.caption(
    "CSV / Excel を放り込むだけで、月次推移・前年比・構成・異常日検知まで自動分析。"
    "列名のゆらぎ(売上/金額/売上高 等)は自動で吸収します。"
)

# ---------- データ入力 ----------
uploaded = st.file_uploader("売上データ(CSV / Excel)をアップロード", type=["csv", "xlsx", "xls"])

if uploaded is not None:
    raw_bytes = uploaded.getvalue()
    filename = uploaded.name
else:
    sample = Path(__file__).parent / "sample-data" / "sales_sample.csv"
    st.info("サンプルデータ(架空の売上・3年分)を表示中。手元のファイルをアップロードすると差し替わります。")
    raw_bytes = sample.read_bytes()
    filename = sample.name

try:
    df_raw = core.read_table(raw_bytes, filename)
    df_norm = core.normalize_columns(df_raw)
    result = core.clean(df_norm)
except core.DataFormatError as e:
    st.error(f"データを読み込めませんでした: {e}")
    st.stop()

df = result.df
if df.empty:
    st.error("有効なデータ行がありません。日付・商品・売上の列と値を確認してください。")
    st.stop()

# ---------- クレンジング報告(黙って捨てない) ----------
if result.dropped:
    detail = " / ".join(f"{k}: {v}行" for k, v in result.dropped.items())
    st.warning(f"⚠️ {result.total_rows:,}行中 {result.kept_rows:,}行を分析対象にしました。除外 → {detail}")
else:
    st.success(f"✅ {result.kept_rows:,}行を読み込みました(除外なし)")

# ---------- KPI ----------
k = core.kpis(df)
c1, c2, c3, c4 = st.columns(4)
c1.metric("総売上", f"¥{k['total']:,.0f}")
c2.metric("月平均", f"¥{k['avg_month']:,.0f}")
c3.metric("最高月", f"¥{k['best_month']:,.0f}")
c4.metric("商品数", f"{k['n_products']}")

# ---------- 月次推移 + 前年比 ----------
st.subheader("月次推移(棒)と前年同月比(線)")
m = core.monthly_summary(df)
fig = go.Figure()
fig.add_bar(x=m["month"], y=m["amount"], name="売上", marker_color="#0f766e")
fig.add_scatter(
    x=m["month"], y=(m["yoy"] * 100).round(1), name="前年比(%)",
    yaxis="y2", mode="lines+markers", line=dict(color="#d97706", width=2),
)
fig.update_layout(
    yaxis=dict(title="売上(円)"),
    yaxis2=dict(title="前年比(%)", overlaying="y", side="right", zeroline=True),
    legend=dict(orientation="h", y=1.12),
    height=420, margin=dict(t=30, b=10),
)
st.plotly_chart(fig, use_container_width=True)

# ---------- 構成 ----------
left, right = st.columns(2)
with left:
    st.subheader("商品別(上位10+その他)")
    bp = core.breakdown(df, "product")
    st.plotly_chart(
        px.bar(bp, x="amount", y="product", orientation="h", color_discrete_sequence=["#0f766e"])
        .update_layout(yaxis=dict(autorange="reversed"), height=380, margin=dict(t=10)),
        use_container_width=True,
    )
with right:
    st.subheader("地域別")
    br = core.breakdown(df, "region")
    st.plotly_chart(
        px.pie(br, values="amount", names="region", hole=0.45)
        .update_layout(height=380, margin=dict(t=10)),
        use_container_width=True,
    )

# ---------- 異常日検知 ----------
st.subheader("異常日検知(日次売上が平均±3σの外側)")
sigma = st.slider("感度(σ)", 2.0, 4.0, 3.0, 0.5, help="小さいほど敏感に検知します")
anom = core.detect_anomalies(df, sigma=sigma)
if anom.empty:
    st.write("異常日は検出されませんでした。")
else:
    show = anom.copy()
    show["amount"] = show["amount"].map(lambda v: f"¥{v:,.0f}")
    show["z"] = show["z"].map(lambda v: f"{v:+.1f}σ")
    show.columns = ["日付", "日次売上", "乖離"]
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.caption("急伸日はキャンペーン効果の確認、急減日は欠品・システム障害の確認を推奨します。")

# ---------- レポート出力 ----------
st.divider()
st.download_button(
    "📥 Excelレポートをダウンロード(月次・商品別・地域別・整形済みデータの4シート)",
    data=core.to_excel_report(df),
    file_name="sales_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.caption("※ サンプルデータは架空の数値です。 SalesLens はポートフォリオ用デモです。")
