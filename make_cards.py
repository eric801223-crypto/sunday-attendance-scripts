# -*- coding: utf-8 -*-
"""
參考優先關心名單｜手機轉發用 PNG 卡片

用法：
    python3 make_cards.py 各區看望名單_YYYYMMDD.txt 卡片輸出資料夾

做法：
只讀取已通過回測的文字名單，不重新計算；每個單位（各區與青職）輸出一張直式 PNG。
姓名前有「＊」者（三份報表模式：近一季小排、家聚會也都沒有紀錄），＊以紅色顯示。
輸出後逐張核對卡片姓名數、註記數與＊數都與文字名單相同，全部相符才顯示「【卡片完成】」。
卡片完成後，再依序號把全部卡片合成一份 PDF（參考優先關心名單_YYYYMMDD.pdf，每個單位一頁），
核對頁數與卡片張數相同才顯示「【PDF完成】」。
"""
import os
import re
import sys
from PIL import Image, ImageDraw, ImageFont

SRC, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)


def font(path, size):
    for i in range(10):
        try:
            f = ImageFont.truetype(path, size, index=i)
        except Exception:
            break
        if "TC" in " ".join(f.getname()):
            return f
    return ImageFont.truetype(path, size)


BOLD = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
MED = "/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc"
REG = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
if not os.path.exists(REG):
    REG = MED

W, PAD = 1080, 60
F_TITLE = font(BOLD, 72)
F_SUB = font(MED, 34)
F_NOTE = font(REG, 26)
F_CAT = font(BOLD, 44)
F_PLAIN = font(MED, 34)
F_COND = font(REG, 28)
F_NAME = font(MED, 54)
F_NAME_S = font(MED, 46)
F_MARK = font(MED, 28)
MARK_H = 40  # 姓名下方「本週有出席」註記所需的額外列高
MARK_COLOR = (46, 125, 50)
STAR = "＊"
STAR_COLOR = (192, 57, 43)

BG = (250, 248, 243)
HEAD = (31, 61, 92)
CARD = (255, 255, 255)
TEXT = (30, 30, 30)
PLAIN = (60, 60, 60)
GREY = (120, 120, 120)
# 順位色：1～2 紅、3～4 橙、5～8 藍綠、9 紫灰
COLORS = {1: (192, 57, 43), 2: (192, 57, 43), 3: (211, 120, 20), 4: (211, 120, 20),
          5: (22, 128, 120), 6: (22, 128, 120), 7: (22, 128, 120), 8: (22, 128, 120), 9: (112, 96, 140)}
RANKS = ["出席穩定者本週缺席", "出席穩定者連續缺席兩週以上", "出席普通者近期缺席", "出席普通者連續缺席三週以上",
         "總出席低，但最近一個月來了一半以上", "總出席低，但最近一個月有來1次",
         "總出席低，原本隔週來，但最近一個月都沒來", "總出席低，原本很少來，且最近一個月都沒來",
         "過去不時出現，但近12週都沒來"]


def wrap(d, text, f, width):
    lines, cur = [], ""
    for ch in text:
        if d.textlength(cur + ch, font=f) > width and ch not in "，。、；：」』）":
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines


def parse(sec):
    lines = sec.strip("\n").splitlines()
    title = lines[0].replace("｜參考優先關心名單", "")
    periods, notes, i = [], [], 1
    while i < len(lines) and lines[i].strip():
        l = lines[i]
        if l.startswith(("統計期間", "第9順位", "本次報表不足一年")):
            periods.append(l)
        else:
            notes.append(l)
        i += 1
    cats, cur, status = [], None, "本次無需關心名單"
    for l in lines[i:]:
        m = re.match(r"^【(.+?)】（(\d+)人）$", l)
        if m:
            cur = dict(name=m.group(1), plain="", cond="", n=int(m.group(2)), people=[], marks={})
            cats.append(cur)
        elif l.startswith("　　") and cur and cur["people"]:
            cur["marks"][len(cur["people"]) - 1] = l.strip("　")
        elif l.startswith("　") and cur:
            cur["people"].append(l[1:])
        elif l.startswith("※") or l.startswith("本次無需"):
            status = l
        elif l.strip() and cur and not cur["people"]:
            if l.startswith("（"):
                cur["cond"] = l
            else:
                cur["plain"] = l
    return title, periods, notes, cats, status


def render(title, periods, notes, cats, status, path):
    tmp = Image.new("RGB", (W, 10))
    d = ImageDraw.Draw(tmp)
    inner = W - 2 * PAD
    cross = "（跨區）" in title  # 青職、大專等跨區名單：兩欄並加註所屬區
    f_name = F_NAME_S if cross else F_NAME
    cols = 2 if cross else 3
    row_h = 78 if cross else 84

    # 量測高度
    note_lines = sum((wrap(d, "※ " + n, F_NOTE, inner - 40) for n in notes), [])
    head_h = 60 + 90 + 50 + 20
    period_lines = sum((wrap(d, p, F_NOTE, inner) for p in periods), [])
    note_h = (len(note_lines) + len(period_lines)) * 40 + 40
    blocks = []
    for c in cats:
        cat_lines = wrap(d, "%s（%d人）" % (c["name"], c["n"]), F_CAT, inner - 80)  # 類別標題過長時自動換行
        plain_lines = wrap(d, c["plain"], F_PLAIN, inner - 80)
        cond_lines = wrap(d, c["cond"], F_COND, inner - 80)
        rows = (len(c["people"]) + cols - 1) // cols
        rh = row_h + (MARK_H if c["marks"] else 0)
        h = 30 + 4 + len(cat_lines) * 60 + len(plain_lines) * 48 + len(cond_lines) * 40 + 16 + rows * rh + 24
        blocks.append((c, cat_lines, plain_lines, cond_lines, rows, h))
    total_people = sum(len(c["people"]) for c in cats)
    H = head_h + 40 + sum(b[5] + 28 for b in blocks) + note_h + 90
    status_lines = wrap(d, status, F_CAT, inner) if not cats else []
    H += len(status_lines) * 64 + 80

    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    # 頁首
    d.rectangle([0, 0, W, head_h], fill=HEAD)
    d.text((PAD, 50), title, font=F_TITLE, fill="white")
    d.text((PAD, 50 + 100), "參考優先關心名單　共%d人" % total_people, font=F_SUB, fill=(220, 230, 240))
    y = head_h + 40

    if not cats:
        warn = status.startswith("※")
        for sl in status_lines:
            d.text((PAD, y + 20), sl, font=F_CAT, fill=(192, 57, 43) if warn else TEXT)
            y += 64
        y += 60

    for c, cat_lines, plain_lines, cond_lines, rows, h in blocks:
        rank = RANKS.index(c["name"]) + 1
        col = COLORS[rank]
        d.rounded_rectangle([PAD, y, W - PAD, y + h], radius=24, fill=CARD, outline=(225, 222, 215), width=2)
        d.rounded_rectangle([PAD, y, PAD + 16, y + h], radius=8, fill=col)
        x0 = PAD + 44
        yy = y + 30
        for cl in cat_lines:
            d.text((x0, yy), cl, font=F_CAT, fill=col)
            yy += 60
        yy += 4
        for pl in plain_lines:
            d.text((x0, yy), pl, font=F_PLAIN, fill=PLAIN)
            yy += 48
        for cl in cond_lines:
            d.text((x0, yy), cl, font=F_COND, fill=GREY)
            yy += 40
        yy += 16
        cw = (W - PAD - 30 - x0) / cols
        rh = row_h + (MARK_H if c["marks"] else 0)
        for i, nm in enumerate(c["people"]):
            r, k = divmod(i, cols)
            fsz = f_name.size
            ff = f_name
            while d.textlength(nm, font=ff) > cw - 12 and fsz > 28:
                fsz -= 2
                ff = font(MED, fsz)
            tx, ty = x0 + k * cw, yy + r * rh + (f_name.size - fsz) * 0.6
            if nm.startswith(STAR):  # 近一季小排、家聚會也都沒有紀錄：＊以紅色標示
                d.text((tx, ty), STAR, font=ff, fill=STAR_COLOR)
                d.text((tx + d.textlength(STAR, font=ff), ty), nm[1:], font=ff, fill=TEXT)
            else:
                d.text((tx, ty), nm, font=ff, fill=TEXT)
            if i in c["marks"]:
                d.text((x0 + k * cw, yy + r * rh + f_name.size + 14), c["marks"][i], font=F_MARK, fill=MARK_COLOR)
        y += h + 28

    # 統計期間與說明
    y += 6
    for nl in period_lines + note_lines:
        d.text((PAD, y), nl, font=F_NOTE, fill=GREY)
        y += 40
    img = img.crop((0, 0, W, y + 50))
    img.save(path)
    return img.size


text = open(SRC, encoding="utf-8").read()
# 文字名單最下面的內部參考（接近穩定名單）不製作卡片、不放入 PDF
text = text.split("\n" + "＝" * 20 + "\n\n")[0]
sections = text.split("\n" + "—" * 20 + "\n\n")
errors, total = [], 0
for i, sec in enumerate(sections, 1):
    title, periods, notes, cats, status = parse(sec)
    n_text = sum(1 for l in sec.splitlines() if l.startswith("　") and not l.startswith("　　"))
    m_text = sum(1 for l in sec.splitlines() if l.startswith("　　"))
    h_text = sum(1 for l in sec.splitlines() if l.startswith("【"))
    m_card = sum(len(c["marks"]) for c in cats)
    s_text = sum(1 for l in sec.splitlines() if l.startswith("　" + STAR))
    s_card = sum(1 for c in cats for nm in c["people"] if nm.startswith(STAR))
    if s_card != s_text:
        errors.append("%s：卡片＊%d筆，文字名單%d筆" % (title, s_card, s_text))
    if h_text != len(cats):
        errors.append("%s：文字名單類別%d個，卡片讀到%d個（類別標題格式可能不符）" % (title, h_text, len(cats)))
    if m_card != m_text:
        errors.append("%s：卡片註記%d筆，文字名單%d筆" % (title, m_card, m_text))
    n_card = sum(len(c["people"]) for c in cats)
    for c in cats:
        if c["n"] != len(c["people"]):
            errors.append("%s【%s】標題%d人，卡片列出%d人" % (title, c["name"], c["n"], len(c["people"])))
        if not c["plain"] or not c["cond"]:
            errors.append("%s【%s】缺少白話說明或數字條件" % (title, c["name"]))
    if n_card != n_text:
        errors.append("%s：卡片%d人，文字名單%d人" % (title, n_card, n_text))
    fn = os.path.join(OUT, "%02d_%s_關心名單.png" % (i, title.replace("（跨區）", "跨區")))
    size = render(title, periods, notes, cats, status, fn)
    total += n_card
    print("%s：%d人，%dx%d" % (os.path.basename(fn), n_card, size[0], size[1]))
if errors:
    print("【卡片失敗】")
    for e in errors:
        print("・" + e)
    sys.exit(1)
print("【卡片完成】共%d張，姓名%d筆，與文字名單相符。" % (len(sections), total))

# 整合 PDF：依卡片序號每張一頁（頁面大小同卡片），只在卡片全部通過後製作
pngs = sorted(f for f in os.listdir(OUT) if re.match(r"^\d{2}_.+_關心名單\.png$", f))
m = re.search(r"(\d{8})", os.path.basename(SRC))
pdf_path = os.path.join(OUT, "參考優先關心名單_%s.pdf" % (m.group(1) if m else "合輯"))
pages = [Image.open(os.path.join(OUT, f)).convert("RGB") for f in pngs]
if len(pages) != len(sections):
    print("【PDF失敗】卡片資料夾有%d張卡片，文字名單有%d個單位。" % (len(pages), len(sections)))
    sys.exit(1)
pages[0].save(pdf_path, save_all=True, append_images=pages[1:], resolution=150)
with open(pdf_path, "rb") as f:
    n_pdf = len(re.findall(rb"/Type\s*/Page[^s]", f.read()))
if n_pdf != len(pages):
    print("【PDF失敗】PDF頁數%d，卡片%d張。" % (n_pdf, len(pages)))
    sys.exit(1)
print("【PDF完成】%s，共%d頁，與卡片張數相同。" % (os.path.basename(pdf_path), n_pdf))
