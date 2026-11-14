#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Notion 頁面 PDF（會所主日指標牧養名單專案共用）

用途：把「寫入 Notion 的頁面內容」原樣排成方便閱讀的 PDF，
關心名單（V7）與近一年分析（V1）共用。不重新計算任何數字，
只排版；排版後從 PDF 抽出文字，逐行核對 Notion 內容都在 PDF 中。

用法：
  python3 -I notion_page_pdf.py 頁面內容.md 輸出.pdf --title "20261111關心名單彙整" \
      [--meta "製表日期：2026-11-11"] [--meta "統計期間：..."]

頁面內容.md：建立 Notion 項目時 content 所用的同一份 Notion Markdown
（## 標題、### 標題、```txt 程式碼區塊、--- 分隔線）。

結果：
  【PDF完成】……   才可傳送 PDF
  【PDF失敗】……   不傳送，把訊息原文告訴使用者
"""
import argparse, html, os, re, shutil, subprocess, sys, tempfile

FS = "　"  # 全形空白


def fail(msg):
    print("【PDF失敗】" + msg)
    sys.exit(1)


# ---------- 解析 Notion Markdown ----------
def parse(md):
    """回傳區塊串列：('h2',text) ('h3',text) ('hr',) ('code',[lines]) ('p',text)"""
    blocks, lines, i = [], md.splitlines(), 0
    while i < len(lines):
        s = lines[i]
        t = s.strip()
        if t.startswith("```"):
            buf, i = [], i + 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i].rstrip("\r"))
                i += 1
            if i >= len(lines):
                fail("程式碼區塊沒有結尾的 ```，請確認頁面內容檔完整。")
            blocks.append(("code", buf))
        elif t == "---":
            blocks.append(("hr",))
        elif t.startswith("### "):
            blocks.append(("h3", t[4:].strip()))
        elif t.startswith("## "):
            blocks.append(("h2", t[3:].strip()))
        elif t.startswith("# "):
            blocks.append(("h2", t[2:].strip()))
        elif t:
            blocks.append(("p", t))
        i += 1
    return blocks


def is_name_line(s):
    return s.startswith(FS) and len(s) > 1 and s[1] != FS


def is_note_line(s):
    return s.startswith(FS + FS)


def is_simple_name(s):
    # 只有姓名（關心名單各區），可含＊與（所屬區）
    body = s[1:]
    return "｜" not in body and len(body) <= 24


# ---------- 排版 ----------
def render_code(lines):
    e = html.escape
    out, i = [], 0
    while i < len(lines):
        s = lines[i]
        if not s.strip():
            out.append('<div class="gap"></div>')
            i += 1
            continue
        # 連續的「只有姓名」行 → 多欄排列（註記行跟在姓名下面）
        if is_name_line(s) and is_simple_name(s):
            cells = []
            while i < len(lines) and is_name_line(lines[i]) and is_simple_name(lines[i]):
                name = lines[i][1:]
                note = ""
                if i + 1 < len(lines) and is_note_line(lines[i + 1]):
                    note = lines[i + 1].strip(FS + " ")
                    i += 1
                star = ""
                if name.startswith("＊"):
                    star, name = '<span class="star">＊</span>', name[1:]
                cells.append(f'<div class="cell"><span class="nm">{star}{e(name)}</span>'
                             + (f'<span class="nt">{e(note)}</span>' if note else "") + "</div>")
                i += 1
            out.append('<div class="grid">' + "".join(cells) + "</div>")
            continue
        if is_name_line(s):
            body = s[1:]
            star = ""
            if body.startswith("＊"):
                star, body = '<span class="star">＊</span>', body[1:]
            m = re.match(r"^(.*?)（([^（）]*｜.*)）$", body)
            if m:
                out.append(f'<div class="pr"><span class="nm">{star}{e(m.group(1))}</span>'
                           f'<span class="dt">（{e(m.group(2))}）</span></div>')
            else:
                out.append(f'<div class="pr"><span class="nm">{star}{e(body)}</span></div>')
        elif is_note_line(s):
            out.append(f'<div class="ind">{e(s.strip(FS))}</div>')
        elif s.startswith("【") or s.startswith("〔"):
            m = re.match(r"^([【〔][^】〕]*[】〕](?:（[^）]*）)?)(.*)$", s)
            rest = m.group(2) if m else ""
            head = m.group(1) if m else s
            out.append(f'<div class="cat"><b>{e(head)}</b>'
                       + (f'<span>{e(rest)}</span>' if rest else "") + "</div>")
        elif s.startswith("・"):
            out.append(f'<div class="bl">{e(s)}</div>')
        elif s.startswith("（") and s.endswith("）"):
            out.append(f'<div class="cond">{e(s)}</div>')
        else:
            out.append(f'<div class="tx">{e(s)}</div>')
        i += 1
    return '<div class="code">' + "".join(out) + "</div>"


CSS = """
@page { size: A4; margin: 15mm 14mm 16mm 14mm; }
body { font-family: 'Noto Sans CJK TC','Noto Sans TC','Microsoft JhengHei',sans-serif;
       font-size: 10.5pt; line-height: 1.6; color:#222; margin:0; }
.title { font-size: 19pt; font-weight: 700; color:#1f3a5f; margin: 0 0 1mm; }
.badge { display:inline-block; background:#b3261e; color:#fff; font-size:9pt; padding:1px 8px;
         border-radius:3px; margin: 0 0 2mm; }
.meta { color:#555; font-size:9.5pt; margin:0; }
h1 { font-size: 14.5pt; color:#fff; background:#1f3a5f; padding:5px 10px; margin: 6mm 0 3mm;
     border-radius:3px; break-after: avoid; }
h2 { font-size: 12.5pt; color:#1f3a5f; border-bottom: 2px solid #c9d6e8; padding-bottom:2px;
     margin: 5mm 0 2mm; break-after: avoid; }
.p { margin: 1mm 0; }
.code { margin: 1mm 0 2mm; }
.gap { height: 2.5mm; }
.tx { margin: .5mm 0; }
.bl { padding-left: 1em; text-indent: -1em; margin: .5mm 0; }
.ind { padding-left: 2.4em; color:#444; font-size:9.5pt; }
.cond { color:#777; font-size:9pt; padding-left: 9px; }
.cat { margin: 3mm 0 1mm; padding: 3px 8px; background:#eef3f9; border-left: 4px solid #1f3a5f;
       break-after: avoid; }
.cat span { color:#555; font-size:9.5pt; margin-left:4px; }
.pr { padding: 1.5px 0 1.5px 1.2em; border-bottom: 1px dotted #ddd; break-inside: avoid; }
.pr .nm { font-weight:700; margin-right: 2px; }
.dt { color:#555; font-size: 9.5pt; }
.grid { display:grid; grid-template-columns: repeat(4, 1fr); column-gap: 4mm; row-gap: 1mm;
        padding: 1mm 0 1mm 1.2em; }
.cell { break-inside: avoid; }
.cell .nm { font-weight:700; font-size: 11pt; display:block; }
.cell .nt { color:#1b7a3a; font-size: 8.5pt; display:block; line-height:1.2; }
.star { color:#b3261e; }
.pb { break-after: page; }
"""


def build_html(blocks, title, metas):
    e = html.escape
    body = []
    for b in blocks:
        k = b[0]
        if k == "h2":
            body.append(f"<h1>{e(b[1])}</h1>")
        elif k == "h3":
            body.append(f"<h2>{e(b[1])}</h2>")
        elif k == "hr":
            body.append('<div class="pb"></div>')
        elif k == "code":
            body.append(render_code(b[1]))
        else:
            body.append(f'<p class="p">{e(b[1])}</p>')
    meta = "".join(f'<p class="meta">{e(m)}</p>' for m in metas)
    return (f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8">'
            f'<title>{e(title)}</title><style>{CSS}</style></head><body>'
            f'<div class="title">{e(title)}</div><div class="badge">內部參考・不對外轉發</div>'
            f'{meta}{"".join(body)}</body></html>')


# ---------- 轉 PDF ----------
def find_chrome():
    for c in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser"):
        p = shutil.which(c)
        if p:
            return p
    return None


def to_pdf(html_path, pdf_path):
    chrome = find_chrome()
    if not chrome:
        fail("工作環境找不到 Chrome／Chromium，無法轉成 PDF。")
    with tempfile.TemporaryDirectory() as prof:
        cmd = [chrome, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
               f"--user-data-dir={prof}", f"--print-to-pdf={pdf_path}", "file://" + html_path]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) < 1000:
        fail("Chrome 沒有產出 PDF。")


def pdf_text(pdf_path):
    if not shutil.which("pdftotext"):
        return None
    r = subprocess.run(["pdftotext", "-enc", "UTF-8", pdf_path, "-"], capture_output=True, timeout=120)
    return r.stdout.decode("utf-8", "replace")


def norm(s):
    return re.sub(r"[\s　]", "", s)


def page_count(pdf_path):
    if shutil.which("pdfinfo"):
        r = subprocess.run(["pdfinfo", pdf_path], capture_output=True, text=True)
        m = re.search(r"Pages:\s+(\d+)", r.stdout)
        if m:
            return int(m.group(1))
    return None


# ---------- 主程式 ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("md")
    ap.add_argument("pdf")
    ap.add_argument("--title", required=True)
    ap.add_argument("--meta", action="append", default=[])
    a = ap.parse_args()

    md = open(a.md, encoding="utf-8").read()
    blocks = parse(md)
    if not any(b[0] == "code" for b in blocks):
        fail("頁面內容檔沒有任何程式碼區塊，可能不是 Notion 頁面內容。")

    # 統計：各 ## 段落的姓名行數（與 Notion 重新讀取後的核對用同一算法）
    sections, cur, total_names = [], None, 0
    for b in blocks:
        if b[0] == "h2":
            cur = [b[1], 0]
            sections.append(cur)
        elif b[0] == "code":
            n = sum(1 for s in b[1] if is_name_line(s))
            total_names += n
            if cur is None:
                cur = ["（開頭）", 0]
                sections.append(cur)
            cur[1] += n

    pdf_path = os.path.abspath(a.pdf)
    with tempfile.TemporaryDirectory() as td:
        hp = os.path.join(td, "page.html")
        open(hp, "w", encoding="utf-8").write(build_html(blocks, a.title, a.meta))
        to_pdf(hp, pdf_path)

    # 核對：Notion 內容每一行（標題、程式碼區塊各行）都出現在 PDF 文字中
    txt = pdf_text(pdf_path)
    if txt is None:
        check = "（工作環境沒有 pdftotext，略過逐行核對）"
    else:
        flat = norm(txt)
        missing = []
        for b in blocks:
            items = b[1] if b[0] == "code" else ([b[1]] if b[0] in ("h2", "h3", "p") else [])
            for s in items:
                key = norm(s)
                if key and key not in flat:
                    # 姓名多欄排列時，姓名與註記分開；只核對姓名與註記各自存在
                    parts = [norm(x) for x in re.split(r"　　", s) if norm(x)]
                    if not all(p in flat for p in parts):
                        missing.append(s.strip())
        # 多欄姓名：PDF 文字順序可能不同，姓名本身也要在
        if missing:
            fail("PDF 中找不到下列 Notion 內容（共%d行）：\n" % len(missing)
                 + "\n".join("  " + m for m in missing[:30]))
        check = "Notion 內容逐行核對通過"

    pages = page_count(pdf_path)
    print("【PDF完成】%s（%s頁）" % (os.path.basename(pdf_path), pages if pages else "?"))
    print("・" + check)
    print("・姓名行合計%d行（各段：%s）" % (total_names, "、".join(f"{t} {n}" for t, n in sections if n)))


if __name__ == "__main__":
    main()
