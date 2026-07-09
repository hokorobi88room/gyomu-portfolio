#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""体験キットZipを personas.json から生成する。
出力: solve/kits/<persona>-kit.zip
各キット: まずこれを読んでください.md / 発注テンプレ.txt / sample-data/ / demo/index.html
"""
import json, os, re, zipfile, io

HERE = os.path.dirname(os.path.abspath(__file__))
SOLVE = os.path.dirname(HERE)
PERSONAS = os.path.join(SOLVE, "personas.json")
TEMPLATE = os.path.join(HERE, "_demo_template.html")
OUT = os.path.join(SOLVE, "kits")

WORK_DESC = {
    "copypaste": "毎朝◯◯のサイトを見て、◯◯をExcelに貼っています",
    "transcribe": "会議やインタビューの録音を、聞いて止めて打っています",
    "invoice": "毎月、台帳を見ながら請求書を1枚ずつ作ってPDFにしています",
    "reception": "申込みが来るたびに、返信して台帳に写して定員を数えています",
    "design": "サムネ（または名札・バナー）を1枚ずつ手作業で作っています",
}
SAMPLE = {
    "copypaste": ("sample-data/収集サンプル.csv",
        "会社名（架空）,業種,所在地\n株式会社アオイ電機,電気設備,東京都\nみどり物流サービス,運送,埼玉県\nハルカ製菓株式会社,食品製造,愛知県\n有限会社ソラ工房,金属加工,大阪府\nつばさ不動産,不動産,福岡県\n"),
    "invoice": ("sample-data/顧客台帳サンプル.csv",
        "顧客コード,取引先名（架空）,締め,支払サイト\n101,株式会社アオイ電機,末,翌月末\n102,みどり物流サービス,末,翌月末\n103,ハルカ製菓株式会社,20日,翌々月10日\n"),
    "reception": ("sample-data/申込サンプル.csv",
        "受付時刻,氏名（架空）,人数,備考\n10:02,佐藤,2,\n10:05,鈴木,1,\n10:11,高橋,3,定員到達\n"),
    "design": ("sample-data/名簿サンプル.csv",
        "氏名（架空）,所属\n田中 太郎,営業部\n佐藤 花子,開発部\n鈴木 一郎,総務部\n高橋 美咲,広報部\n"),
    "transcribe": ("sample-data/サンプルについて.txt",
        "この体験版では、実際の音声ファイルは同梱していません。\n"
        "デモは「文字が起きていく様子」を再現したものです（内容は架空）。\n"
        "本番では、お預かりした録音（会議・インタビュー等）を自動でテキスト化し、\n"
        "話者や見出しで整えたたたき台をお渡しします。秘密保持を前提に進めます。\n"),
}


def strip_h1(h1):
    return re.sub(r"<[^>]+>", " ", h1).replace("——", "").strip()


def readme(pid, p):
    r = p["result"]
    sol = r["solution"]
    price = r.get("price", "内容によります")
    lines = []
    lines.append("# " + strip_h1(r["h1"]))
    lines.append("")
    lines.append("> これは「体験版」です。実際に動きますが、途中で止まります。")
    lines.append("> 続きは、あなたの本物のデータでお見せします。")
    lines.append("")
    lines.append("## まず、触ってみてください")
    lines.append("`demo/index.html` をダブルクリックで開き、ボタンを押してください。")
    lines.append("ブラウザだけで動きます（インストール不要）。サンプルは架空です。")
    lines.append("")
    lines.append("## これは、こういうことです")
    lines.append("- **" + sol["claim"] + "**")
    lines.append("- " + sol["reason"])
    lines.append("")
    lines.append("## あなたのデータで動かすには")
    lines.append("この体験版はサンプル専用です。あなたの本物の作業で動かすには、")
    lines.append("いまの作業内容を送っていただくだけで大丈夫です。")
    lines.append("同梱の **発注テンプレ.txt** をコピーして送れば、それだけで相談がはじまります。")
    lines.append("")
    lines.append("- 当方の目安：**" + price + "**（" + r.get("priceNote", "") + "）")
    lines.append("- ご提案・お見積りまで無料です。お返事は24時間以内を心がけています。")
    lines.append("- ご相談：https://coconala.com/users/6171856")
    lines.append("")
    lines.append("---")
    lines.append("サンプルの会社名・数値はすべて架空です。／ 綻 流夢（ほころび るうむ）")
    return "\n".join(lines) + "\n"


def order_template(pid):
    work = WORK_DESC.get(pid, "毎回やっている、決まった手作業があります")
    return (
        "【コピペで使える依頼テンプレ】そのまま送って大丈夫です。\n"
        "----------------------------------------\n"
        "はじめまして。診断ページ（体験キット）から来ました。\n\n"
        "・いま手作業でやっていること:\n"
        "  （例）" + work + "\n\n"
        "・頻度と時間:\n"
        "  （例）週に◯回、1回◯分くらい\n\n"
        "・最後どうなっていれば最高か:\n"
        "  （例）ボタン1つ、または開くだけで終わっている状態\n\n"
        "まずは実現できそうか、概算とあわせて教えてください。\n"
        "----------------------------------------\n"
        "そのまま貼り付けて、◯の部分だけ埋めれば送れます。\n"
    )


def demo_html(template, pid, p):
    r = p["result"]
    demo = r["demo"]
    h1 = strip_h1(r["h1"])
    cfg = dict(demo)
    html = template.replace("__TITLE__", h1)
    html = html.replace("__H1__", h1)
    html = html.replace("__BTN__", demo.get("buttonLabel", "▶ 実行する"))
    html = html.replace("__CONFIG__", json.dumps(cfg, ensure_ascii=False))
    return html


def main():
    data = json.load(open(PERSONAS, encoding="utf-8"))
    template = open(TEMPLATE, encoding="utf-8").read()
    os.makedirs(OUT, exist_ok=True)
    built = []
    for pid, p in data["personas"].items():
        if not p.get("kit"):
            continue
        base = pid + "-kit"
        files = {
            base + "/まずこれを読んでください.md": readme(pid, p),
            base + "/発注テンプレ.txt": order_template(pid),
            base + "/demo/index.html": demo_html(template, pid, p),
        }
        sname, sbody = SAMPLE[pid]
        files[base + "/" + sname] = sbody
        zpath = os.path.join(OUT, base + ".zip")
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            for name, body in files.items():
                zi = zipfile.ZipInfo(name)
                zi.flag_bits |= 0x800  # UTF-8 filename
                z.writestr(zi, body.encode("utf-8"))
        built.append((base + ".zip", len(files)))
    for name, n in built:
        print("built", name, "(%d files)" % n)


if __name__ == "__main__":
    main()
