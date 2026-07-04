"""各作品フォルダに、非エンジニア向けの起動ファイル(HTML)と取扱説明書を生成する。

macOSの警告(Gatekeeper)を避けるため、.command/.bat は使わず HTML を採用。
HTMLファイルはダブルクリックすると警告なしでブラウザが開く(Mac/Windows共通)。
"""
from __future__ import annotations

from pathlib import Path

WORKS = Path(__file__).resolve().parent.parent / "works"

OLD = ["▶Macで開く.command", "▶Windowsで開く.bat", "はじめにお読みください.txt",
       "run_mac.command", "run_windows.bat"]

KIND_REAL = {
    "vba": "※これはExcelのマクロ(VBA)です。実際の業務で使う本物は、あなたのExcelに組み込んでボタン1つで動かします(Windowsが確実)。この体験版はブラウザで仕組みを見るものです。",
    "gas": "※本物はGoogleフォーム＋スプレッドシートに組み込み、24時間自動で動きます。この体験版は流れを見るものです。",
    "python": "※本物はPCやクラウドで動くPythonツールです。クラウドに置けばURLを開くだけでも使えます。",
    "go": "※本物はインストール不要の単一ファイル(アプリ)で、この体験版と同じ処理を一瞬で行います。",
    "web": "※これはそのまま公開できるWebページ(HTML)です。実案件ではあなたの文章・写真に差し替えます。",
}

WORKS_DATA = {
    "a1-invoice-vba": ("vba", "請求書 自動発行システム",
        "顧客リストと売上から、請求書を全件まとめてPDFにするExcelツールです。",
        ["ボタン1回で全顧客分の請求書PDFを作成", "インボイス(税率別)・源泉徴収に対応", "手作業で半日 → 数分に短縮"], "demo/index.html"),
    "a2-report-vba": ("vba", "複数ブック 自動集計ダッシュボード",
        "様式がバラバラな月次報告ファイル(20個)を、1つに自動でまとめるExcelツールです。",
        ["フォルダに入れて実行するだけで自動統合", "列の並びや呼び方の違いを自動で吸収", "前年比・ランキングまで自動作成"], "demo/index.html"),
    "a3-form-gas": ("gas", "フォーム受付の 完全自動化",
        "申込フォームの受付を、確認メール・台帳・通知・定員管理まで全部自動にする仕組みです。",
        ["申込と同時に確認メールを自動返信", "台帳へ自動記録・重複もチェック", "定員に達したら自動で受付終了"], "demo/index.html"),
    "w1-course-lp": ("web", "高コンバージョンLP",
        "広告の受け皿になる『売れる縦長ページ』の制作サンプルです。",
        ["売れる型で構成を設計", "各セクションの意図を見せる解説モードつき", "スマホ最優先・高速表示"], "index.html"),
    "w2-cafe-site": ("web", "店舗ホームページ(4ページ)",
        "小さなお店向けの、きれいで必要十分なホームページ(4ページ)のサンプルです。",
        ["トップ/メニュー/店舗案内/お知らせ", "地図・営業時間まで情報設計込み", "検索対策(SEO)にも対応"], "index.html"),
    "p1-sales-dashboard": ("python", "売上分析ダッシュボード",
        "売上データ(CSV)を入れるだけで自動分析するツールです。",
        ["前年比・商品別を自動でグラフ化", "売上が急に増減した日を自動で発見", "結果をExcelに出力"], "demo/index.html"),
    "p2-webwatch-excel": ("python", "Web情報の定点観測ツール",
        "複数サイトの公開情報を自動で見に行き『昨日から変わった所』だけを教えるツールです。",
        ["毎朝の巡回をコマンド1回に", "増えた行=緑・消えた行=赤で色分け", "規約を守った行儀の良い収集"], "demo/index.html"),
    "g1-gocrunch": ("go", "大量ファイル 高速処理ツール",
        "Excelでは開けない数千〜数万のファイルを、一瞬で結合・集計するツールです。",
        ["2000ファイルの結合を実測0.024秒", "列の並び違いを自動でそろえる", "インストール不要の単一ファイル"], "demo/index.html"),
}

LAUNCH_HTML = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8">
<meta http-equiv="refresh" content="0; url={target}">
<title>デモを開いています…</title>
<style>body{{font-family:sans-serif;background:#12203c;color:#fff;text-align:center;padding:80px 20px}}a{{color:#5eead4;font-size:1.2rem}}</style>
</head><body>
<p style="font-size:1.3rem">デモを開いています…</p>
<p>自動で開かない場合は <a href="{target}">こちらをクリック</a></p>
</body></html>
"""

TXT = """====================================================
  {title}
====================================================

【これは何？】
{overview}

【ここがすごい】
{sugoi}

----------------------------------------------------
【使い方 — かんたん2ステップ】
----------------------------------------------------
1. ダウンロードしたZIPファイルをダブルクリックして解凍します
   （フォルダが出てきます）

2. フォルダの中の
   「▶デモを見る（ダブルクリック）.html」
   をダブルクリックします。

   → インターネットの画面（ブラウザ）でデモが開きます。
     あとは画面の青いボタンを押すだけ！ すぐに動きます。

★ ダブルクリックするだけ。インストールも設定も不要、
  こわい警告も出ません。安心してお試しください。

----------------------------------------------------
{real}

このデモのデータ・社名などはすべて架空のサンプルです。
ご依頼をいただければ、あなたの実際の業務に合わせてお作りします。
"""


def main() -> None:
    for wid, (kind, title, overview, sugoi, target) in WORKS_DATA.items():
        d = WORKS / wid
        if not d.exists():
            continue
        for old in OLD:
            p = d / old
            if p.exists():
                p.unlink()
        sugoi_txt = "\n".join(f"  ・{s}" for s in sugoi)
        (d / "▶デモを見る（ダブルクリック）.html").write_text(
            LAUNCH_HTML.format(target=target), encoding="utf-8")
        (d / "まずはこちらをお読みください（取扱説明書）.txt").write_text(
            TXT.format(title=title, overview=overview, sugoi=sugoi_txt, real=KIND_REAL[kind]),
            encoding="utf-8")
        print("ok ->", wid)


if __name__ == "__main__":
    main()
