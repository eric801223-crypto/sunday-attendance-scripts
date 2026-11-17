# -*- coding: utf-8 -*-
"""把「關心名單互動版」與「近一年分析互動版」合成一個 HTML，兩個大分頁各自獨立。

每個分頁是一個獨立的內嵌頁面（iframe srcdoc），樣式、篩選、展開狀態互不影響；
內容原封不動取自兩個已產出的 HTML，不重新計算。第二個分頁第一次切換時才載入，確保圖表寬度正確。

用法：python3 -I combine_html.py 關心名單互動版.html 近一年出席分析互動版.html 輸出.html [--made 製表日期]
"""
import argparse
import html
import json

ap = argparse.ArgumentParser()
ap.add_argument('care')
ap.add_argument('annual')
ap.add_argument('out')
ap.add_argument('--made', default='')
a = ap.parse_args()

pages = []
for path, label in ((a.care, '近12週關心名單'), (a.annual, '近一年分析')):
    with open(path, encoding='utf-8') as f:
        doc = f.read()
    if '/*DATA*/' in doc or 'const D = ' not in doc:
        raise SystemExit(f'【停止】{path} 不是已產出的互動版 HTML。')
    pages.append((label, doc))

tabs = ''.join(f'<button class="t{" on" if i == 0 else ""}" data-i="{i}">{html.escape(l)}</button>' for i, (l, _) in enumerate(pages))
frames = ''.join(f'<iframe id="f{i}" title="{html.escape(l)}"{"" if i == 0 else " hidden"}></iframe>' for i, (l, _) in enumerate(pages))
docs = json.dumps([d for _, d in pages], ensure_ascii=False).replace('</', '<\\/')

out = f'''<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>會所主日出席｜互動版{(" " + html.escape(a.made)) if a.made else ""}</title>
<style>
:root{{color-scheme:light;--bg:#f6f5f2;--bar:#fcfcfb;--border:#e4e2dc;--text:#0b0b0b;--muted:#8a8984;--warn-bg:#fdecec;--warn:#b42318}}
@media (prefers-color-scheme: dark){{:root{{color-scheme:dark;--bg:#121211;--bar:#1a1a19;--border:#2e2e2b;--text:#f0efec;--muted:#8f8e86;--warn-bg:#3a1d1b;--warn:#ff9b8f}}}}
html,body{{margin:0;height:100%;background:var(--bg);font-family:"PingFang TC","Noto Sans TC","Microsoft JhengHei",system-ui,sans-serif}}
.top{{height:48px;display:flex;align-items:center;gap:6px;padding:0 14px;background:var(--bar);border-bottom:1px solid var(--border);box-sizing:border-box}}
.t{{border:1px solid var(--border);background:transparent;color:var(--text);font-family:inherit;font-size:15px;padding:6px 16px;border-radius:999px;cursor:pointer}}
.t.on{{background:var(--text);color:var(--bg);border-color:var(--text);font-weight:700}}
.sp{{flex:1}}
.b{{background:var(--warn-bg);color:var(--warn);font-size:12px;font-weight:700;padding:2px 10px;border-radius:999px;white-space:nowrap}}
.m{{color:var(--muted);font-size:12.5px;white-space:nowrap}}
@media (max-width:560px){{.m{{display:none}}.t{{padding:6px 11px;font-size:14px}}}}
iframe{{display:block;width:100%;height:calc(100% - 48px);border:0}}
iframe[hidden]{{display:none}}
</style>
</head>
<body>
<div class="top">{tabs}<span class="sp"></span><span class="m">{html.escape(("製表日期：" + a.made) if a.made else "")}</span><span class="b">內部參考，不對外轉發</span></div>
{frames}
<script>
const DOCS = {docs};
const loaded = new Set();
function show(i) {{
  document.querySelectorAll('.t').forEach(b => b.classList.toggle('on', +b.dataset.i === i));
  document.querySelectorAll('iframe').forEach((f, k) => f.hidden = k !== i);
  if (!loaded.has(i)) {{ document.getElementById('f' + i).srcdoc = DOCS[i]; loaded.add(i); }}
  try {{ history.replaceState(null, '', '#' + i); }} catch (e) {{}}  // 預覽視窗等環境不允許改網址時略過
}}
document.querySelectorAll('.t').forEach(b => b.onclick = () => show(+b.dataset.i));
let start = 0; try {{ if (/^#[01]$/.test(location.hash)) start = +location.hash.slice(1); }} catch (e) {{}}
show(start);
</script>
</body>
</html>
'''
with open(a.out, 'w', encoding='utf-8') as f:
    f.write(out)
print(f'【HTML完成】{a.out}（兩個分頁：{"、".join(l for l, _ in pages)}）')
