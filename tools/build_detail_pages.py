"""作品詳細ページ(README → HTML)とソース一式zipの生成。

Netlify等の静的ホスティング単体で完結させるため、GitHubリンクの
代わりに works/<id>/index.html(詳細ページ)と <id>.zip(DL用)を作る。
再実行可能(冪等): 生成物を作り直すだけ。
"""
from __future__ import annotations

import io
import zipfile
from pathlib import Path

import markdown

ROOT = Path(__file__).resolve().parent.parent
WORKS = ROOT / "works"

# 詳細ページを生成する作品(W1/W2はそれ自体がHTML作品なので対象外)
TARGETS = {
    "a1-invoice-vba": {"demo": "demo/index.html", "setup": "docs/SETUP.md"},
    "a2-report-vba": {"demo": "demo/index.html", "setup": "docs/SETUP.md"},
    "a3-form-gas": {"demo": "demo/index.html", "setup": "docs/SETUP.md"},
    "p1-sales-dashboard": {"demo": "demo/index.html"},
    "p2-webwatch-excel": {"demo": "demo/index.html"},
    "g1-gocrunch": {"demo": "demo/index.html"},
}

TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} | 業務自動化×Web制作 ポートフォリオ</title>
<meta name="description" content="{title} — 作品詳細・ソースコード">
<style>
  :root{{--bg:#f6f8fb;--card:#fff;--ink:#16202e;--sub:#5a6a80;--line:#e2e8f1;--teal:#0f766e;--teal2:#0d9488;--navy:#12203c}}
  @media (prefers-color-scheme: dark){{:root{{--bg:#0d1420;--card:#161f2e;--ink:#e9eef6;--sub:#94a4ba;--line:#273349;--navy:#0a1226}}}}
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:"Noto Sans JP","Hiragino Sans",sans-serif;background:var(--bg);color:var(--ink);line-height:1.9}}
  .top{{background:var(--navy);color:#fff;padding:20px 0}}
  .wrap{{max-width:860px;margin:0 auto;padding:0 22px}}
  .top a{{color:#5eead4;text-decoration:none;font-size:.85rem;font-weight:700}}
  .top h1{{font-size:1.3rem;margin-top:6px}}
  .actions{{display:flex;gap:12px;flex-wrap:wrap;margin:26px 0 8px}}
  .btn{{display:inline-block;text-decoration:none;font-weight:700;border-radius:10px;padding:11px 24px;font-size:.9rem}}
  .btn-main{{background:var(--teal);color:#fff}}
  .btn-main:hover{{background:var(--teal2)}}
  .btn-sub{{border:1px solid var(--line);color:var(--ink);background:var(--card)}}
  .btn-sub:hover{{border-color:var(--teal2);color:var(--teal2)}}
  article{{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:34px 36px;margin:18px 0 60px}}
  article h1{{font-size:1.45rem;margin-bottom:18px}}
  article h2{{font-size:1.12rem;margin:30px 0 12px;padding-left:12px;border-left:4px solid var(--teal)}}
  article p{{margin:10px 0;font-size:.93rem}}
  article ul,article ol{{margin:10px 0 10px 1.5em;font-size:.93rem}}
  article table{{border-collapse:collapse;width:100%;margin:14px 0;font-size:.87rem}}
  article th,article td{{border:1px solid var(--line);padding:9px 12px;text-align:left}}
  article th{{background:rgba(13,148,136,.1)}}
  article code{{background:rgba(13,148,136,.12);border-radius:4px;padding:1px 6px;font-size:.85em}}
  article pre{{background:#0b1220;color:#c9f0dd;border-radius:10px;padding:16px;overflow-x:auto;margin:14px 0}}
  article pre code{{background:none;padding:0;color:inherit}}
  article a{{color:var(--teal2)}}
  article hr{{border:none;border-top:1px dashed var(--line);margin:24px 0}}
  article em{{color:var(--sub)}}
</style>
</head>
<body>
<div class="top"><div class="wrap"><a href="../../index.html">← ポートフォリオへ戻る</a><h1>{title}</h1></div></div>
<div class="wrap">
  <div class="actions">{buttons}</div>
  <article>{body}</article>
</div>
</body>
</html>
"""


def md_to_html(md_path: Path) -> tuple[str, str]:
    """README/SETUPをHTML化してタイトル(h1)と本文を返す。"""
    text = md_path.read_text(encoding="utf-8")
    title = next(
        (line.lstrip("# ").strip() for line in text.splitlines() if line.startswith("# ")),
        md_path.stem,
    )
    body = markdown.markdown(text, extensions=["tables", "fenced_code"])
    return title, body


def make_zip(work_dir: Path) -> Path:
    """作品フォルダのソース一式zip(生成物のzip/詳細HTML自身は除外)。"""
    zip_path = work_dir / f"{work_dir.name}.zip"
    exclude = {zip_path.name, "index.html", "setup.html"}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(work_dir.rglob("*")):
            if f.is_dir():
                continue
            rel = f.relative_to(work_dir)
            # 生成物(zip・詳細HTML)はzipに含めない。demo/やランチャは含める
            if str(rel) in exclude:
                continue
            arcname = f"{work_dir.name}/{rel}"
            data = f.read_bytes()
            info = zipfile.ZipInfo(arcname)
            info.compress_type = zipfile.ZIP_DEFLATED
            # .command / .sh は実行ビットを保持(解凍後もダブルクリックで起動できるように)
            mode = 0o755 if f.suffix in (".command", ".sh") else 0o644
            info.external_attr = mode << 16
            zf.writestr(info, data)
    zip_path.write_bytes(buf.getvalue())
    return zip_path


def main() -> None:
    for work_id, opts in TARGETS.items():
        work_dir = WORKS / work_id
        readme = work_dir / "README.md"
        if not readme.exists():
            raise SystemExit(f"README がありません: {work_dir}")

        zip_path = make_zip(work_dir)

        buttons = [f'<a class="btn btn-main" href="{work_id}.zip" download>📥 ソース一式をダウンロード</a>']
        if demo := opts.get("demo"):
            buttons.insert(0, f'<a class="btn btn-main" href="{demo}">▶ ブラウザで体験デモ</a>')
        if setup := opts.get("setup"):
            s_title, s_body = md_to_html(work_dir / setup)
            (work_dir / "setup.html").write_text(
                TEMPLATE.format(title=s_title, buttons=f'<a class="btn btn-sub" href="index.html">← 作品詳細へ戻る</a>', body=s_body),
                encoding="utf-8",
            )
            buttons.append('<a class="btn btn-sub" href="setup.html">📖 導入手順を見る</a>')

        title, body = md_to_html(readme)
        (work_dir / "index.html").write_text(
            TEMPLATE.format(title=title, buttons="".join(buttons), body=body),
            encoding="utf-8",
        )
        print(f"built {work_id}: index.html"
              + (" setup.html" if "setup" in opts else "")
              + f" {zip_path.name} ({zip_path.stat().st_size // 1024}KB)")


if __name__ == "__main__":
    main()
