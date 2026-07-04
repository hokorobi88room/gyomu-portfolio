"""各作品の添付用サムネ画像(cover.png)を生成する。

SVGを組み立て、rsvg-convert で PNG(1200x750) に変換して
各 works/<id>/cover.png に出力する。CWのポートフォリオ添付用。
"""
from __future__ import annotations

import subprocess
from pathlib import Path

WORKS = Path(__file__).resolve().parent.parent / "works"

FONT = "'Hiragino Sans','Hiragino Kaku Gothic ProN','Arial Unicode MS','Noto Sans JP',sans-serif"

# id, タグ, タグ色, タイトル行(最大2), サブ
ITEMS = [
    ("a1-invoice-vba", "EXCEL VBA", "#1d6f42", ["請求書を", "全件自動発行"], "月4時間 → 3分。PDF一括・インボイス対応"),
    ("a2-report-vba", "EXCEL VBA", "#1d6f42", ["複数ファイルを", "自動集計"], "様式バラバラでも1つに統合・前年比まで"),
    ("a3-form-gas", "GAS", "#3f6212", ["フォーム受付を", "完全自動化"], "自動返信・台帳転記・通知・定員クローズ"),
    ("w1-course-lp", "WEB / LP", "#1e40af", ["売れる", "縦長LP"], "構成の意図まで見せる解説モードつき"),
    ("w2-cafe-site", "WEB / HP", "#1e40af", ["店舗ホーム", "ページ（4P）"], "営業時間・メニュー・地図まで情報設計"),
    ("p1-sales-dashboard", "PYTHON", "#0e7490", ["売上分析", "ダッシュボード"], "CSVを入れるだけ・異常値検知・レポート出力"),
    ("p2-webwatch-excel", "PYTHON", "#0e7490", ["Web情報の", "定点観測ツール"], "自動収集→差分をExcel化・規約順守"),
    ("g1-gocrunch", "GO", "#0369a1", ["大量ファイル", "高速処理CLI"], "2000ファイル結合 0.024秒（実測）"),
]

W, H = 1200, 750


def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def svg(tag, tagcolor, title_lines, sub) -> str:
    # タイトル行
    ty = 300 if len(title_lines) == 2 else 340
    title_tspans = ""
    for i, line in enumerate(title_lines):
        title_tspans += f'<text x="90" y="{ty + i*108}" font-size="94" font-weight="800" fill="#ffffff" font-family="{FONT}">{esc(line)}</text>\n'
    return f'''<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">
  <defs>
    <radialGradient id="glow" cx="0.82" cy="0.12" r="0.6">
      <stop offset="0" stop-color="#0d9488" stop-opacity="0.55" />
      <stop offset="1" stop-color="#0d9488" stop-opacity="0" />
    </radialGradient>
  </defs>
  <rect width="{W}" height="{H}" fill="#12203c" />
  <rect width="{W}" height="{H}" fill="url(#glow)" />
  <rect x="0" y="0" width="14" height="{H}" fill="#d9a441" />
  <rect x="90" y="120" rx="22" ry="22" width="{60 + len(tag)*24}" height="46" fill="{tagcolor}" />
  <text x="{90 + 24}" y="151" font-size="26" font-weight="700" fill="#ffffff" letter-spacing="3" font-family="{FONT}">{esc(tag)}</text>
  {title_tspans}
  <text x="92" y="{ty + len(title_lines)*108 + 20}" font-size="34" fill="#f0c260" font-family="{FONT}">{esc(sub)}</text>
  <text x="90" y="{H-52}" font-size="26" fill="#93a3b8" font-family="{FONT}">業務自動化 × Web制作 ｜ 実務6年 × AI</text>
</svg>'''


def main() -> None:
    scratch = Path("/private/tmp/claude-501/-Users-kmacmini-Desktop-Fable5/16092b33-7a71-4656-9d43-4e7e344ba68e/scratchpad")
    for wid, tag, tagcolor, title_lines, sub in ITEMS:
        s = svg(tag, tagcolor, title_lines, sub)
        svg_path = scratch / f"{wid}.svg"
        svg_path.write_text(s, encoding="utf-8")
        out = WORKS / wid / "cover.png"
        subprocess.run(
            ["rsvg-convert", "-w", str(W), "-h", str(H), "-o", str(out), str(svg_path)],
            check=True,
        )
        print(f"cover.png -> {wid}  ({out.stat().st_size // 1024}KB)")


if __name__ == "__main__":
    main()
