# -*- coding: utf-8 -*-
"""
主日出席分析與關懷優先序｜通用規則 V7 分析程式

用法：
    只有主日：
    python3 sunday_attendance.py 主日報表.xlsx 輸出資料夾 [--zero-week include|exclude]
    主日＋小排＋家聚會：
    python3 sunday_attendance.py 主日報表.xlsx 輸出資料夾 --group 小排報表.xlsx --home 家聚會報表.xlsx [--zero-week ...]

說明：
- 以報表內「製表日期」為基準日，不依賴系統時間。
- 主日日期＝報表欄位日期（週一）＋6天；只採用主日日期早於製表日期的週次。
- 主日已過但全區出席為0的週次：預設停止並說明，需以 --zero-week 指定處理方式。
- 三份報表模式：分類與排序仍只依主日；列入名單者若近一季（同統計期間12週）小排、家聚會都沒有紀錄，姓名前加「＊」。
  另列出主日被排除、但近一季小排或家聚會仍有紀錄者（「主日沒來但有小排家聚會」內部名單：Excel 工作表與文字名單最下面，
  不製作卡片，2026-11-12 起）。
- 跨區名單：群組「青職」產出「青職（跨區）」、群組「大學」產出「大專（跨區）」（2026-11-11 新增），
  姓名後加註所屬區，不重複計入全區人數。
"""
import re
import sys
import argparse
import datetime as dt
from collections import defaultdict, OrderedDict, Counter

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

AREAS = ["1小區", "2小區", "3小區", "4小區", "5小區", "6小區", "7小區", "青年區", "兒童專項"]
EXCLUDE_KEYWORDS = ["來訪", "訪客", "臨時", "福音朋友"]
YEAR_DAYS = 364
YEAR_MIN = 6  # 第9順位門檻：近12週0次，前一年內出席6次以上
MARK_CATS = (5, 6)  # 這兩類若最近一次主日有來，在姓名下方註記
MARK_TEXT = "（本週有出席）"
STAR = "＊"
STAR_NOTE = "＊表示近一季（同統計期間12週）小排、家聚會也都沒有紀錄"
# 跨區名單：（群組欄的值、文字名單標題、Excel 工作表名稱）；不重複計入全區人數
CROSS = [("青職", "青職（跨區）", "青職"), ("大學", "大專（跨區）", "大專")]
# 內部參考：接近穩定名單（2026-11-12 新增）。只放 Excel 工作表與文字名單檔最下面，不進卡片、PDF、Notion
INTERNAL_SEP = "\n" + "＝" * 20 + "\n\n"
NEAR_TITLE = "接近穩定名單｜內部參考，不對外轉發"
NEAR_SHEET = "接近穩定（內部）"
NEAR_NOTE = ("說明：12週來6～8次，且沒有連續缺席3週以上（含這週有來者及最近一兩週沒來者），"
             "再多來一兩次就達到出席穩定。用途是鼓勵「從常來變成每週來」，只供負責弟兄參考。")
# 內部參考：主日沒來但有小排家聚會（2026-11-12 起獨立成名單；只在三份報表模式，不製作卡片、不放入 PDF）
CONTACT_TITLE = "主日沒來但有小排家聚會｜內部參考，不對外轉發"
CONTACT_SHEET = "主日沒來但有小排家聚會（內部）"
CONTACT_NOTE = ("說明：近12週主日都沒來、近一年內主日也不到%d次（因此不在關心名單上），"
                "但近一季（同統計期間12週）小排或家聚會有紀錄。仍有連結，可邀請來主日。只供負責弟兄參考。" % YEAR_MIN)


def near_people(res):
    """接近穩定：12週6～8次且連續缺席未達3週（即出席普通（本週有來）與第3順位），含兒童專項，不含登載示警的區。"""
    return [p for p in res["people"] if p["cat"] in ("N", 3) and p["area"] not in res["warn_areas"]]


def near_key(p):
    order = {a: i for i, a in enumerate(AREAS)}
    return (order[p["area"]], -p["n"], p["k"], p["name"])


def near_line(p):
    return "　%s（%s｜12週%d次｜%s）" % (p["name"], p["group"] or "無群組", p["n"], "本週有來" if p["k"] == 0 else "最近缺席%d週" % p["k"])


def near_export(res):
    ps = sorted(near_people(res), key=near_key)
    lines = [NEAR_TITLE] + period_lines(res)[:1] + [NEAR_NOTE, "共%d人" % len(ps), ""]
    for a in AREAS:
        lst = [p for p in ps if p["area"] == a]
        if not lst:
            continue
        lines.append("【%s】（%d人）" % (a, len(lst)))
        lines += [near_line(p) for p in lst]
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def contact_line(e):
    return "　%s（%s｜小排%d週、家聚會%d週｜近一個月小排%d週、家聚會%d週）" % (
        e["name"], e["group"] or "無群組", e["g12"], e["h12"], e["g4"], e["h4"])


def contact_export(res):
    ps = res["excl_contact"]
    lines = [CONTACT_TITLE] + period_lines(res)[:1] + [CONTACT_NOTE, "共%d人" % len(ps), ""]
    for a in AREAS:
        lst = [e for e in ps if e["area"] == a]
        if not lst:
            continue
        lines.append("【%s】（%d人）" % (a, len(lst)))
        lines += [contact_line(e) for e in lst]
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"

NOTE_LINES = [
    "穩定無狀況、近12週0出席且近一年內出席未達6次不納入（可能為長期停用未處理名單），請自行掌握",
    "名單逕以點名系統匯出分析，名單可能包含「已遷出但在條件範圍內曾出席」、「應停用、遷出但尚未於系統調整」名單。",
]

# 順位：（類別名稱、白話說明、數字條件）
CATS = OrderedDict([
    (1, ("出席穩定者本週缺席", "往常固定來，這週沒來", "12週來9次以上，上週有來、這週沒來")),
    (2, ("出席穩定者連續缺席兩週以上", "往常固定來，但連續幾週沒來", "12週來9次以上，已連續兩週以上沒來")),
    (3, ("出席普通者近期缺席", "往常還算常來，最近一兩週沒來", "12週來6～8次，最近一兩週沒來")),
    (4, ("出席普通者連續缺席三週以上", "往常還算常來，但已有一段時間沒來", "12週來6～8次，已連續三週以上沒來")),
    (5, ("總出席低，但最近一個月來了一半以上", "原本不算常來或是新人，最近一個月來了一半以上", "前8週來0～3次，近4週來2次以上；12週合計5次以下")),
    (6, ("總出席低，但最近一個月有來1次", "原本不常來或是新人，最近一個月來了1次", "前8週來0～4次，近4週來1次")),
    (7, ("總出席低，原本隔週來，但最近一個月都沒來", "原本大約隔週會來，但最近一個月都沒來", "前8週來3～5次，近4週0次")),
    (8, ("總出席低，原本很少來，且最近一個月都沒來", "原本很少來，最近一個月也都沒來", "前8週來1～2次，近4週0次")),
    (9, ("過去不時出現，但近12週都沒來", "以前偶爾會出現，但近三個月都沒見到", "近12週0次，近一年內來過6次以上")),
])
NOT_LISTED = OrderedDict([
    ("S", ("出席穩定（本週有來）", "12週來9次以上，這週有來")),
    ("N", ("出席普通（本週有來）", "12週來6～8次，這週有來")),
])


def fail(msg):
    print("【停止】" + msg)
    sys.exit(2)


def read_report(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["出席"] if "出席" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))

    ref_date = None
    header_idx = None
    for i, r in enumerate(rows[:10]):
        for v in r:
            if isinstance(v, str):
                m = re.search(r"製表日期\s*[:：]\s*(\d{4})-(\d{1,2})-(\d{1,2})", v)
                if m:
                    ref_date = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if r and "姓名" in [str(x).strip() if x is not None else "" for x in r]:
            header_idx = i
            break
    if ref_date is None:
        fail("報表中找不到「製表日期」，無法判定哪些主日已舉行。")
    if header_idx is None:
        fail("報表中找不到含「姓名」的標題列。")

    header = [str(x).strip() if x is not None else "" for x in rows[header_idx]]
    col = {name: header.index(name) for name in ["姓名", "大區", "小區", "群組"] if name in header}
    for need in ["姓名", "大區", "小區"]:
        if need not in col:
            fail("標題列缺少「%s」欄位。" % need)
    date_cols = []
    for j, h in enumerate(header):
        m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", h)
        if m:
            date_cols.append((j, dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))))
    data = [r for r in rows[header_idx + 1:] if r and r[col["姓名"]]]
    return ref_date, col, date_cols, data


def pick_weeks(ref_date, date_cols, data, zero_week):
    """回傳：最近12次主日、其前一年內的主日（用於第9順位）、週次說明、是否涵蓋一年。"""
    past, notes = [], []
    # 同一週出現兩欄以上（例如跨年週）：只有一欄有出席資料就採用該欄；都有資料則停止請使用者確認
    by_mon = OrderedDict()
    for j, mon in date_cols:
        by_mon.setdefault(mon, []).append(j)
    dedup = []
    for mon, js in by_mon.items():
        if len(js) == 1:
            dedup.append((js[0], mon))
            continue
        sunday = mon + dt.timedelta(days=6)
        filled = [j for j in js if sum(int(r[j] or 0) for r in data) > 0]
        if len(filled) > 1:
            fail("%s 主日在報表中有%d個重複週次欄，且都有出席資料，無法判斷採用哪一欄。請確認點名系統後重新匯出報表。"
                 % (sunday.isoformat(), len(js)))
        keep = filled[0] if filled else js[0]
        dedup.append((keep, mon))
        if filled:
            notes.append("%s 主日報表有重複週次欄（%d欄），採用有出席資料的那一欄，其餘空白欄不納入。" % (sunday.isoformat(), len(js)))
    date_cols = dedup
    for j, mon in date_cols:
        sunday = mon + dt.timedelta(days=6)
        if sunday >= ref_date:
            notes.append("%s 主日尚未舉行（製表日期 %s），不納入。" % (sunday.isoformat(), ref_date.isoformat()))
            continue
        past.append((j, sunday, sum(int(r[j] or 0) for r in data)))
    if not past:
        fail("報表中沒有已舉行的主日。")
    last = past[-1][1]
    year_start = last - dt.timedelta(days=YEAR_DAYS)
    used = [x for x in past if x[1] > year_start]
    # 近12週以前的週次只用於第9順位：全區出席為0者直接視為未點名、不納入，不需詢問
    nonzero = [x for x in used if x[2] > 0]
    cutoff = nonzero[-12][1] if len(nonzero) >= 12 else used[0][1]
    for j, sunday, total in used:
        if total == 0 and sunday < cutoff:
            notes.append("%s 主日全區出席為0（較近12週更早，視為當週未點名），不納入。" % sunday.isoformat())
    used = [x for x in used if not (x[2] == 0 and x[1] < cutoff)]
    held = []
    for j, sunday, total in used:
        if total == 0:
            if zero_week is None:
                fail("%s 主日日期已過，但全區出席人數為0。可能尚未登載或當週停聚會。"
                     "請確認後以 --zero-week include（計入，全體視為缺席）或 --zero-week exclude（不計入）重新執行。"
                     % sunday.isoformat())
            if zero_week == "exclude":
                notes.append("%s 主日全區出席為0，依指示不納入。" % sunday.isoformat())
                continue
            notes.append("%s 主日全區出席為0，依指示計入。" % sunday.isoformat())
        held.append((j, sunday))
    if len(held) < 12:
        fail("已舉行的主日只有 %d 次，不足12次。" % len(held))
    weeks = held[-12:]
    earlier = held[:-12]
    year_ok = past[0][1] <= year_start + dt.timedelta(days=7)
    if not year_ok:
        notes.append("報表期間不足一年（最早主日 %s），第9順位無法完整判定；請匯出一年期報表。" % past[0][1].isoformat())
    return weeks, earlier, notes, year_ok


def area_of(r, col):
    big = (r[col["大區"]] or "").strip()
    small = (r[col["小區"]] or "").strip() if r[col["小區"]] else ""
    issue = None
    if big == "會所青年區":
        a = "青年區"
    elif big == "兒童專項":
        a = "兒童專項"
        if small == "青年區":
            issue = "大區為兒童專項、小區填青年區，依規則歸兒童專項"
    else:
        a = small
        if a not in AREAS:
            issue = "無法判定分區（大區：%s，小區：%s）" % (big or "空白", small or "空白")
    return a, issue


def classify(seq):
    n = sum(seq)
    k = 0
    for x in reversed(seq):
        if x == 0:
            k += 1
        else:
            break
    r4 = sum(seq[-4:])
    if n >= 9:
        cat = "S" if k == 0 else (1 if k == 1 else 2)
    elif n >= 6:
        cat = "N" if k == 0 else (3 if k <= 2 else 4)
    else:
        # 總出席低（1～5次）：先看近4週，近4週0次者再依前8週（即12週）出席次數區分
        cat = 5 if r4 >= 2 else (6 if r4 == 1 else (7 if n >= 3 else 8))
    return n, k, r4, cat


def sort_key(p):
    c = p["cat"]
    last = p["last"] or dt.date(1900, 1, 1)
    if c == 1:
        return (-p["n"], p["name"])
    if c in (2, 3, 4):
        return (p["k"], -p["n"], p["name"])
    if c in (5, 6):
        return (-p["r4"], -last.toordinal(), p["name"])
    if c == 9:
        return (-p["prior"], -last.toordinal(), p["name"])
    return (-last.toordinal(), p["name"])


def row_key(r, col):
    """比對三份報表用：姓名＋大區＋小區（同區同名者再依名單順序第幾筆區分）。"""
    return (str(r[col["姓名"]]).strip(), str(r[col["大區"]] or "").strip(), str(r[col["小區"]] or "").strip())


def load_aux(path, kind, ref_date, mons, sunday_total):
    """讀取小排或家聚會報表，回傳 {(姓名, 大區, 小區, 第幾筆): [12週各週出席]} 與每個姓名鍵的筆數。"""
    aref, acol, adates, adata = read_report(path)
    if aref != ref_date:
        fail("%s報表製表日期 %s 與主日報表 %s 不同，請匯出同一天的報表。" % (kind, aref, ref_date))
    have = set(d for _, d in adates)
    missing = [m for m in mons if m not in have]
    if missing:
        fail("%s報表缺少統計期間的週次：%s" % (kind, "、".join(m.isoformat() for m in missing)))
    recs, cnt = {}, Counter()
    for r in adata:
        k = row_key(r, acol)
        cnt[k] += 1
        vals = []
        for m in mons:
            vals.append(max(int(r[j] or 0) for j, d in adates if d == m))
        recs[k + (cnt[k],)] = vals
    total = sum(sum(v) for v in recs.values())
    if total >= sunday_total:
        fail("%s報表近12週出席總數（%d）不少於主日報表（%d），檔案可能放錯，請確認哪一份是主日報表。" % (kind, total, sunday_total))
    return recs, cnt


def build(path, zero_week=None, group=None, home=None):
    ref_date, col, date_cols, data = read_report(path)
    weeks, earlier, week_notes, year_ok = pick_weeks(ref_date, date_cols, data, zero_week)
    sundays = [s for _, s in weeks]

    people, excluded, issues = [], defaultdict(int), []
    excl_zero = []  # 近12週0出席且一年內未達門檻者（用於三份報表模式的接觸比對）
    scnt = Counter()
    for r in data:
        name = str(r[col["姓名"]]).strip()
        rk = row_key(r, col)
        scnt[rk] += 1
        key = rk + (scnt[rk],)
        texts = " ".join(str(r[col[c]] or "") for c in col)
        a, issue = area_of(r, col)
        if any(kw in texts for kw in EXCLUDE_KEYWORDS):
            excluded[(a, "臨時或來訪")] += 1
            continue
        seq = [int(r[j] or 0) for j, _ in weeks]
        prior = [int(r[j] or 0) for j, _ in earlier]
        if sum(seq) == 0 and sum(prior) < YEAR_MIN:
            excluded[(a, "近12週0出席且一年內未達%d次" % YEAR_MIN)] += 1
            excl_zero.append(dict(name=name, area=a, key=key, group=(r[col["群組"]] if "群組" in col else "") or ""))
            continue
        if issue:
            issues.append("%s：%s" % (name, issue))
        if a not in AREAS:
            continue
        n, k, r4, cat = classify(seq)
        last = max((sundays[i] for i, x in enumerate(seq) if x), default=None)
        if n == 0:
            cat = 9
            k = 12 + next((i for i, x in enumerate(reversed(prior)) if x), len(prior))
            last = max(earlier[i][1] for i, x in enumerate(prior) if x)
        people.append(dict(name=name, area=a, group=(r[col["群組"]] if "群組" in col else "") or "",
                           seq=seq, n=n, k=k, r4=r4, cat=cat, last=last, prior=sum(prior), key=key, star=False))

    # 三份報表模式：小排、家聚會接觸比對
    aux = bool(group or home)
    excl_contact = []
    if aux:
        if not (group and home):
            fail("小排、家聚會報表需兩份一起提供（--group 與 --home）。")
        mons = [s - dt.timedelta(days=6) for s in sundays]
        sunday_total = sum(int(r[j] or 0) for r in data for j, _ in weeks)
        g_recs, g_cnt = load_aux(group, "小排", ref_date, mons, sunday_total)
        h_recs, h_cnt = load_aux(home, "家聚會", ref_date, mons, sunday_total)

        def contact(p):
            rk = p["key"][:3]
            if g_cnt[rk] == 0 and h_cnt[rk] == 0:
                return None, "notfound"
            if g_cnt[rk] != scnt[rk] or h_cnt[rk] != scnt[rk]:
                return None, "mismatch"
            g = g_recs[p["key"]]
            h = h_recs[p["key"]]
            return dict(g12=sum(1 for x in g if x), h12=sum(1 for x in h if x),
                        g4=sum(1 for x in g[-4:] if x), h4=sum(1 for x in h[-4:] if x)), None

        for p in people:
            c, err = contact(p)
            if err == "notfound":
                p.update(g12=0, h12=0, g4=0, h4=0)
                if p["cat"] in CATS:
                    issues.append("%s（%s）：小排、家聚會報表中找不到此人（姓名或分區可能不一致），視為近一季無紀錄並標＊" % (p["name"], p["area"]))
            elif err == "mismatch":
                p.update(g12=None, h12=None, g4=None, h4=None)
                if p["cat"] in CATS:
                    issues.append("%s（%s）：同區同名筆數在三份報表中不一致，無法比對小排、家聚會，不標＊" % (p["name"], p["area"]))
            else:
                p.update(c)
            p["star"] = p["cat"] in CATS and p["g12"] == 0 and p["h12"] == 0
        for e in excl_zero:
            c, err = contact(e)
            if c and (c["g12"] or c["h12"]):
                e.update(c)
                excl_contact.append(e)
        order = {a: i for i, a in enumerate(AREAS)}
        excl_contact.sort(key=lambda e: (order.get(e["area"], 99), -(e["g4"] + e["h4"]), -(e["g12"] + e["h12"]), e["name"]))

    # 登載示警
    warn_areas = []
    for a in AREAS:
        ps = [p for p in people if p["area"] == a]
        if not ps:
            continue
        last_total = sum(p["seq"][-1] for p in ps)
        prev_avg = sum(sum(p["seq"][:-1]) for p in ps) / 11
        if last_total == 0 and prev_avg > 0:
            warn_areas.append(a)

    return dict(ref_date=ref_date, sundays=sundays, week_notes=week_notes, people=people, year_ok=year_ok,
                year_from=(earlier[0][1] if earlier else None),
                excluded=excluded, issues=issues, warn_areas=warn_areas, total_rows=len(data),
                aux=aux, excl_contact=excl_contact)


def disp(p, show_area=False):
    return (STAR if p.get("star") else "") + p["name"] + ("（%s）" % p["area"] if show_area else "")


def note_lines(res):
    return NOTE_LINES + ([STAR_NOTE] if res["aux"] else [])


def marked(p):
    return p["cat"] in MARK_CATS and p["seq"][-1] > 0


def cat_label(c):
    return CATS[c][0] if c in CATS else NOT_LISTED[c][0]


def period_lines(res):
    s = res["sundays"]
    first = "統計期間：最近12次主日（%s～%s）" % (s[0].strftime("%Y/%m/%d"), s[-1].strftime("%Y/%m/%d"))
    if res["year_ok"] and res["year_from"]:
        second = "第9順位另參考近一年紀錄（%s～%s）" % (res["year_from"].strftime("%Y/%m/%d"), s[-1].strftime("%Y/%m/%d"))
    else:
        second = "本次報表不足一年，第9順位未完整判定"
    return [first, second]


def text_export(res, area, people, show_area=False):
    s = res["sundays"]
    lines = ["%s｜參考優先關心名單" % area,
             ] + period_lines(res) + note_lines(res) + [""]
    if area in res["warn_areas"]:
        lines += ["※ 本區最近一次主日出席為0，可能尚未登載，暫不產出看望名單。", ""]
        return "\n".join(lines)
    shown = 0
    for c, (nm, plain, cond) in CATS.items():
        ps = sorted([p for p in people if p["cat"] == c], key=sort_key)
        if not ps:
            continue
        shown += 1
        lines.append("【%s】（%d人）" % (nm, len(ps)))
        lines.append(plain)
        lines.append("（%s）" % cond)
        for p in ps:
            lines.append("　" + disp(p, show_area))
            if marked(p):
                lines.append("　　" + MARK_TEXT)
        lines.append("")
    if shown == 0:
        lines += ["本次無需關心名單。", ""]
    return "\n".join(lines).rstrip() + "\n"


def summary_rows(res, people_by_area):
    rows = []
    for a, ps in people_by_area.items():
        cnt = {c: sum(1 for p in ps if p["cat"] == c) for c in list(CATS) + list(NOT_LISTED)}
        stars = [sum(1 for p in ps if p["star"])] if res["aux"] else []
        rows.append([a, len(ps)] + [cnt[c] for c in CATS] +
                    [cnt["S"], cnt["N"]] + stars + ["可能尚未登載" if a in res["warn_areas"] else ""])
    return rows


# ---------- 重點摘要（2026-11-13 新增） ----------
HL_TITLE = "關心名單重點摘要｜內部參考，不對外轉發"
HL_INVITE = 2  # 主日沒來但有小排家聚會：近一個月小排加家聚會2週以上，列為可邀回主日


def hl_lists(res):
    """重點摘要的名單（回測程式獨立重算核對）。"""
    order = {a: i for i, a in enumerate(AREAS)}
    ps = [p for p in res["people"] if p["area"] not in res["warn_areas"]]
    key = lambda p: (p["cat"], order[p["area"]], sort_key(p))
    urgent = sorted([p for p in ps if p["cat"] in (1, 2)], key=key)
    fading = sorted([p for p in ps if p["cat"] == 4 or (res["aux"] and p["cat"] == 7 and p["star"])], key=key)
    lost = sorted([p for p in ps if p["cat"] == 9 and p["star"]], key=key) if res["aux"] else []
    invite = [e for e in res["excl_contact"] if e["g4"] + e["h4"] >= HL_INVITE] if res["aux"] else []
    return urgent, fading, lost, invite


def highlights_text(res):
    ps = [p for p in res["people"] if p["area"] not in res["warn_areas"]]
    listed = [p for p in ps if p["cat"] in CATS]
    L = [HL_TITLE, "製表日期：%s" % res["ref_date"], period_lines(res)[0],
         "說明：依固定規則從關心名單挑出建議先留意的人與地方，只是提出問題，不代表原因，也不據以評價區或個人。名單可能包含已遷出或應停用但系統尚未調整的人，發現時請自行略過。", ""]

    L.append("【一、本週概況】")
    c = lambda lo, hi: sum(1 for p in listed if lo <= p["cat"] <= hi)
    L.append("・列入關心名單共%d人：原本常來的人缺席（第1～4順位）%d人、原本不常來（第5～8順位）%d人、近三個月都沒見到（第9順位）%d人。" % (
        len(listed), c(1, 4), c(5, 8), c(9, 9)))
    by = Counter(p["area"] for p in listed if p["cat"] <= 4)
    L.append("・第1～4順位依區：" + ("、".join("%s %d人" % (a, by[a]) for a in sorted(by, key=lambda a: (-by[a], AREAS.index(a)))) or "無") + "。")
    if res["aux"]:
        st = [p for p in listed if p["star"]]
        L.append("・列入名單者中，小排、家聚會近一季也都沒有紀錄（＊）%d人，其中第1～4順位%d人、第9順位%d人。" % (
            len(st), sum(1 for p in st if p["cat"] <= 4), sum(1 for p in st if p["cat"] == 9)))
    L.append("")

    urgent, fading, lost, invite = hl_lists(res)

    def line(p):
        if p["cat"] == 9:
            mid = "近12週都沒來，之前一年內來過%d次" % p["prior"]
        else:
            mid = "12週來%d次，連續缺席%d週" % (p["n"], p["k"])
        star = "，小排家聚會也沒有紀錄" if p["star"] else ""
        return "　%s%s（%s｜第%d順位，%s%s）" % (STAR if p["star"] else "", p["name"], p["area"], p["cat"], mid, star)

    L.append("【二、建議本週優先聯絡】")
    L.append("原本固定來（12週來9次以上），這週或連續幾週沒來（第1、2順位）；越早聯絡越容易接住。" +
             ("姓名前有＊者，小排、家聚會近一季也沒有紀錄，請更優先。" if res["aux"] else ""))
    L.append("〔共%d人〕" % len(urgent))
    L += [line(p) for p in urgent] or ["　無"]
    L.append("")

    L.append("【三、可能正在淡出】")
    L.append("原本還算常來，已連續三週以上沒來（第4順位）" +
             ("；以及原本隔週來、近一個月都沒來，且小排、家聚會也沒有紀錄（第7順位＊）" if res["aux"] else "") + "。")
    L.append("已經一段時間沒來，再晚可能就接不住了；比第5～8順位更需要主動聯絡。")
    L.append("〔共%d人〕" % len(fading))
    L += [line(p) for p in fading] or ["　無"]
    if res["aux"]:
        L.append("〔可能已斷了聯繫：近三個月都沒見到，小排、家聚會也沒有紀錄（第9順位＊）〕（%d人）從第9順位中挑出主日之外也沒有接觸的人，可能需要特別去找。" % len(lost))
        L += [line(p) for p in lost] or ["　無"]
    L.append("")

    L.append("【四、正在回來的人】")
    back = [p for p in ps if p["cat"] in MARK_CATS and p["seq"][-1] > 0]
    bb = Counter(p["area"] for p in back)
    L.append("・原本不常來或新人，這週有來（第5、6順位，名單上標「本週有出席」）%d人%s；建議見面時表達歡迎、持續邀約。" % (
        len(back), "（" + "、".join("%s %d人" % (a, bb[a]) for a in AREAS if bb[a]) + "）" if back else ""))
    L.append("・接近穩定（12週來6～8次，再多來一兩次就達到穩定）%d人，名單見文字名單最下面；可鼓勵從常來變成每週來。" % len(near_people(res)))
    L.append("")

    if res["aux"]:
        L.append("【五、可以邀回主日】")
        L.append("不在關心名單上（近12週主日都沒來），但近一個月小排或家聚會來了%d週以上；仍有連結，建議透過小排、家聚會邀請來主日。" % HL_INVITE)
        L.append("從文字名單最下面「主日沒來但有小排家聚會」的%d人中，只列近一個月來%d週以上的人：最近還常接觸，邀請最自然。" % (len(res["excl_contact"]), HL_INVITE))
        L.append("〔共%d人〕" % len(invite))
        L += ["　%s（%s｜近一個月小排%d週、家聚會%d週）" % (e["name"], e["area"], e["g4"], e["h4"]) for e in invite] or ["　無"]
        L.append("")

    L.append("【資料提醒】")
    tips = []
    if res["warn_areas"]:
        tips.append("・%s 最近一次主日出席為0，可能尚未登載，本次暫不產出名單，請確認後重跑。" % "、".join(res["warn_areas"]))
    if res["issues"]:
        tips.append("・資料待確認%d筆（見 Excel「排除紀錄及資料待確認」）。" % len(res["issues"]))
    if not res["year_ok"]:
        tips.append("・報表不足一年，第9順位未完整判定。")
    L += tips or ["・無。"]
    return "\n".join(L).rstrip() + "\n"

def write_excel(res, out_path, highlights=""):
    people = res["people"]
    s = res["sundays"]
    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="BBBBBB")
    hdr_font = Font(bold=True, size=12)
    hdr_fill = PatternFill("solid", fgColor="E7F1FB")

    def table(ws, start_row, headers, rows, widths=None):
        for j, h in enumerate(headers, 1):
            c = ws.cell(row=start_row, column=j, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        for i, row in enumerate(rows, start_row + 1):
            for j, v in enumerate(row, 1):
                c = ws.cell(row=i, column=j, value=v)
                c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                c.font = Font(size=12)
        if widths:
            for j, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(j)].width = w
        return start_row + len(rows) + 1

    by_area = OrderedDict((a, [p for p in people if p["area"] == a]) for a in AREAS)

    # 1 分類概況
    ws = wb.active
    ws.title = "分類概況"
    ws["A1"] = "主日出席分類概況"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = "統計期間：%s～%s（最近12次主日）　製表日期：%s" % (s[0], s[-1], res["ref_date"])
    extra = ["標＊人數"] if res["aux"] else []
    headers = ["分區", "有效名單"] + ["第%d順位" % c for c in CATS] + ["出席穩定（本週有來）", "出席普通（本週有來）"] + extra + ["登載示警"]
    rows = summary_rows(res, by_area)
    tot = ["合計", sum(r[1] for r in rows)] + [sum(r[2 + i] for r in rows) for i in range(len(CATS) + 2 + len(extra))] + [""]
    end = table(ws, 4, headers, rows + [tot], [10, 10] + [9] * len(CATS) + [12, 12] + [10] * len(extra) + [14])
    for c in range(1, len(headers) + 1):
        ws.cell(row=end, column=c).font = Font(bold=True, size=12)
    r = end + 2
    ws.cell(row=r, column=1, value="備註：各欄意義").font = Font(bold=True, size=12)
    legend = ["有效名單：點名報表全部名單，扣除姓名或欄位含「%s」的臨時點名對象，以及近12週0出席且近一年內出席未達%d次者（可能為長期停用未處理名單）"
              % ("」「".join(EXCLUDE_KEYWORDS), YEAR_MIN)]
    legend += ["第%d順位：%s｜%s（%s）" % (c, nm, plain, cond) for c, (nm, plain, cond) in CATS.items()]
    legend += ["%s：%s，不列入看望順位" % (nm, cond) for nm, cond in NOT_LISTED.values()]
    if res["aux"]:
        legend += ["標＊人數：列入第1～9順位者中，近一季（同統計期間12週）小排、家聚會也都沒有紀錄的人數（姓名前加＊）",
                   "主日被排除、但近一季小排或家聚會仍有紀錄者，見「%s」工作表（%d人），另列於文字名單最下面，不列入對外名單" % (CONTACT_SHEET, len(res["excl_contact"]))]
    for t in legend:
        r += 1
        ws.cell(row=r, column=1, value=t).font = Font(size=12)

    # 2 分析說明
    ws = wb.create_sheet("分析說明")
    notes = ["依《主日出席分析與關懷優先序｜通用規則 V7》分析。",
             "分析模式：" + ("主日＋小排＋家聚會（分類只依主日；近一季小排、家聚會都沒有紀錄者姓名前加＊）" if res["aux"] else "只有主日"),
             "製表日期：%s（以報表內製表日期為基準，不依賴系統時間）" % res["ref_date"],
             "採用的12次主日：" + "、".join(d.strftime("%m/%d") for d in s),
             ] + period_lines(res)
    notes += ["未採用週次：" + n for n in res["week_notes"]]
    notes += ["", "看望順位："]
    notes += ["%d. %s：%s（%s）" % (c, nm, plain, cond) for c, (nm, plain, cond) in CATS.items()]
    notes += ["不列入看望順位：出席穩定（12週來9次以上，這週有來）、出席普通（12週來6～8次，這週有來）",
              "", "排序原則：先看平常出席程度，同一程度內，最近越有出席的人越優先。"]
    for i, t in enumerate(notes, 1):
        ws.cell(row=i, column=1, value=t).font = Font(size=12)
    ws.column_dimensions["A"].width = 100

    # 3 各區名單（含青職跨區）
    cross_title = {sheet: title for _, title, sheet in CROSS}
    pages = list(by_area.items()) + [(sheet, [p for p in people if p["group"] == g and p["area"] not in res["warn_areas"]])
                                     for g, _, sheet in CROSS]
    for a, ps in pages:
        show_area = a in cross_title
        ws = wb.create_sheet(a)
        ws.column_dimensions["A"].width = 70
        r = 1
        ws.cell(row=r, column=1, value="%s｜參考優先關心名單" % (cross_title[a] if show_area else a)).font = Font(bold=True, size=18)
        r += 1
        for t in period_lines(res) + note_lines(res):
            ws.cell(row=r, column=1, value=t).font = Font(size=12)
            r += 1
        r += 1
        if a in res["warn_areas"]:
            ws.cell(row=r, column=1, value="※ 本區最近一次主日出席為0，可能尚未登載，暫不產出看望名單。").font = Font(size=14, color="C00000")
            continue
        for c, (nm, plain, cond) in CATS.items():
            lst = sorted([p for p in ps if p["cat"] == c], key=sort_key)
            if not lst:
                continue
            cell = ws.cell(row=r, column=1, value="【%s】（%d人）" % (nm, len(lst)))
            cell.font = Font(bold=True, size=15)
            cell.fill = hdr_fill
            r += 1
            ws.cell(row=r, column=1, value=plain).font = Font(size=13, color="333333")
            r += 1
            ws.cell(row=r, column=1, value="（%s）" % cond).font = Font(size=12, color="777777")
            r += 1
            if lst:
                for i, p in enumerate(lst, 1):
                    ws.cell(row=r, column=1, value="%d. %s" % (i, disp(p, show_area))).font = Font(size=15)
                    r += 1
                    if marked(p):
                        ws.cell(row=r, column=1, value="　　" + MARK_TEXT).font = Font(size=12, color="2E7D32")
                        r += 1
            else:
                ws.cell(row=r, column=1, value="無名單").font = Font(size=15, color="888888")
                r += 1
            r += 1

    # 4 全區看望總表
    ws = wb.create_sheet("全區看望總表")
    aux_h = ["小排近一季", "家聚會近一季", "小排近一個月", "家聚會近一個月", "標＊"] if res["aux"] else []

    def aux_v(p):
        if not res["aux"]:
            return []
        f = lambda v: "無法比對" if v is None else v
        return [f(p["g12"]), f(p["h12"]), f(p["g4"]), f(p["h4"]), STAR if p["star"] else ""]

    headers = ["順位", "類別", "所屬區", "姓名", "最近出席日期", "連續缺席週數", "近四週出席", "12週出席", "近12週以前一年內出席", "最近12次出席紀錄（左舊右新）"] + aux_h
    rows = []
    for c, (nm, *_) in CATS.items():
        for p in sorted([p for p in people if p["cat"] == c and p["area"] not in res["warn_areas"]], key=sort_key):
            rows.append([c, nm, p["area"], p["name"], p["last"].isoformat() if p["last"] else "",
                         p["k"], p["r4"], p["n"], p["prior"], "".join("●" if x else "○" for x in p["seq"])] + aux_v(p))
    table(ws, 1, headers, rows, [6, 34, 10, 22, 14, 12, 10, 10, 14, 30] + [10] * len(aux_h))
    ws.freeze_panes = "A2"

    # 5 跨區分類概況 / 6 跨區完整名單（青職、大專）
    for g, _, sheet in CROSS:
        yp = [p for p in people if p["group"] == g]
        ws = wb.create_sheet("%s分類概況" % sheet)
        headers = ["類別", "人數"]
        rows = [["第%d順位｜%s" % (c, nm), sum(1 for p in yp if p["cat"] == c)] for c, (nm, *_) in CATS.items()]
        rows += [["不列入｜%s" % nm, sum(1 for p in yp if p["cat"] == k)] for k, (nm, _) in NOT_LISTED.items()]
        rows.append(["合計", len(yp)])
        table(ws, 1, headers, rows, [46, 10])
        ws = wb.create_sheet("%s跨區完整名單" % sheet)
        headers = ["順位", "類別", "所屬區", "姓名", "最近出席日期", "連續缺席週數", "近四週出席", "12週出席"] + aux_h
        rows = []
        for c in list(CATS) + list(NOT_LISTED):
            for p in sorted([p for p in yp if p["cat"] == c], key=sort_key if c in CATS else (lambda p: (-p["n"], p["name"]))):
                rows.append([c if c in CATS else "不列入", cat_label(c), p["area"], p["name"],
                             p["last"].isoformat() if p["last"] else "", p["k"], p["r4"], p["n"]] + aux_v(p))
        table(ws, 1, headers, rows, [8, 34, 10, 22, 14, 12, 10, 10] + [10] * len(aux_h))
        ws.freeze_panes = "A2"

    # 7 12次出席矩陣
    ws = wb.create_sheet("12次出席矩陣")
    headers = ["所屬區", "姓名", "群組"] + [d.strftime("%m/%d") for d in s] + ["12週出席", "出席率", "連續缺席週數", "近四週出席", "近12週以前一年內出席", "最近出席日期", "類別"] + aux_h
    rows = []
    for a in AREAS:
        for p in sorted([p for p in people if p["area"] == a], key=lambda p: p["name"]):
            rows.append([a, p["name"], p["group"]] + p["seq"] +
                        [p["n"], p["n"] / 12, p["k"], p["r4"], p["prior"], p["last"].isoformat() if p["last"] else "",
                         ("第%d順位｜" % p["cat"] if p["cat"] in CATS else "不列入｜") + cat_label(p["cat"])] + aux_v(p))
    end = table(ws, 1, headers, rows, [10, 22, 8] + [7] * 12 + [9, 8, 12, 10, 14, 13, 40] + [10] * len(aux_h))
    for r in range(2, end):
        ws.cell(row=r, column=17).number_format = "0%"
    ws.freeze_panes = "D2"

    # 7-1 主日沒來但有小排家聚會（只在三份報表模式；內部參考）
    if res["aux"]:
        ws = wb.create_sheet(CONTACT_SHEET)
        ws["A1"] = CONTACT_TITLE
        ws["A1"].font = Font(bold=True, size=13)
        ws["A2"] = CONTACT_NOTE
        ws["A2"].font = Font(size=12)
        headers = ["所屬區", "姓名", "群組", "小排近一季", "家聚會近一季", "小排近一個月", "家聚會近一個月"]
        rows = [[e["area"], e["name"], e["group"], e["g12"], e["h12"], e["g4"], e["h4"]] for e in res["excl_contact"]]
        table(ws, 3, headers, rows, [10, 22, 8, 12, 12, 12, 12])
        ws.freeze_panes = "A4"

    # 7-2 接近穩定（內部參考）
    ws = wb.create_sheet(NEAR_SHEET)
    ws["A1"] = NEAR_TITLE
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = NEAR_NOTE
    ws["A2"].font = Font(size=12)
    headers = ["所屬區", "姓名", "群組", "12週出席", "近四週出席", "連續缺席週數", "目前情形", "最近12次出席紀錄（左舊右新）"]
    rows = [[p["area"], p["name"], p["group"], p["n"], p["r4"], p["k"],
             "本週有來" if p["k"] == 0 else "最近缺席（第3順位）", "".join("●" if x else "○" for x in p["seq"])]
            for p in sorted(near_people(res), key=near_key)]
    table(ws, 4, headers, rows, [10, 22, 8, 10, 10, 12, 18, 30])
    ws.freeze_panes = "A5"

    # 8 排除紀錄及資料待確認
    ws = wb.create_sheet("排除紀錄及資料待確認")
    headers = ["分區", "排除原因", "人數"]
    rows = [[a, why, v] for (a, why), v in sorted(res["excluded"].items(), key=lambda x: (str(x[0][0]), x[0][1]))]
    end = table(ws, 1, headers, rows, [14, 20, 8])
    r = end + 1
    ws.cell(row=r, column=1, value="資料待確認").font = Font(bold=True, size=13)
    for t in res["issues"] or ["無"]:
        r += 1
        ws.cell(row=r, column=1, value=t).font = Font(size=12)
    r += 2
    ws.cell(row=r, column=1, value="週次判定").font = Font(bold=True, size=13)
    for t in res["week_notes"] or ["無"]:
        r += 1
        ws.cell(row=r, column=1, value=t).font = Font(size=12)

    if highlights:
        ws = wb.create_sheet("重點摘要", 0)
        for i, t in enumerate(highlights.rstrip("\n").split("\n"), 1):
            ws.cell(row=i, column=1, value=t).font = Font(bold=t.startswith("【") or i == 1, size=14 if i == 1 else 12)
        ws.column_dimensions["A"].width = 120
        wb.active = 0
    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("outdir")
    ap.add_argument("--zero-week", choices=["include", "exclude"])
    ap.add_argument("--group", help="小排報表")
    ap.add_argument("--home", help="家聚會報表")
    args = ap.parse_args()

    import os
    os.makedirs(args.outdir, exist_ok=True)
    res = build(args.report, args.zero_week, args.group, args.home)
    people = res["people"]
    s = res["sundays"]
    tag = s[-1].strftime("%Y%m%d")

    # 完整性檢查
    n_valid = len(people)
    n_ex = sum(res["excluded"].values())
    unassigned = res["total_rows"] - n_valid - n_ex
    cat_sum = sum(1 for p in people if p["cat"] in CATS) + sum(1 for p in people if p["cat"] in NOT_LISTED)

    hl = highlights_text(res)
    with open(os.path.join(args.outdir, "關心名單重點摘要_%s.txt" % tag), "w", encoding="utf-8") as f:
        f.write(hl)
    write_excel(res, os.path.join(args.outdir, "主日出席分析_%s.xlsx" % tag), hl)
    all_txt = []
    for a in AREAS:
        t = text_export(res, a, [p for p in people if p["area"] == a])
        all_txt.append(t)
    cross_txt = {}
    for g, title, _ in CROSS:
        cross_txt[g] = [p for p in people if p["group"] == g and p["area"] not in res["warn_areas"]]
        all_txt.append(text_export(res, title, cross_txt[g], show_area=True))
    with open(os.path.join(args.outdir, "各區看望名單_%s.txt" % tag), "w", encoding="utf-8") as f:
        f.write(("\n" + "—" * 20 + "\n\n").join(all_txt) + INTERNAL_SEP + near_export(res) +
                (INTERNAL_SEP + contact_export(res) if res["aux"] else ""))

    print("分析模式：%s" % ("主日＋小排＋家聚會" if res["aux"] else "只有主日"))
    print("製表日期：%s" % res["ref_date"])
    print("採用12次主日：%s" % "、".join(d.isoformat() for d in s))
    for n in res["week_notes"]:
        print("未採用：" + n)
    if not res["year_ok"]:
        print("【注意】報表期間不足一年，第9順位無法完整判定，請匯出一年期報表。")
    print("報表總人數：%d；排除：%d；有效名單：%d；無法分區：%d" % (res["total_rows"], n_ex, n_valid, unassigned))
    for c, (nm, *_) in CATS.items():
        print("第%d順位 %s：%d" % (c, nm, sum(1 for p in people if p["cat"] == c)))
    for k, (nm, _) in NOT_LISTED.items():
        print("不列入 %s：%d" % (nm, sum(1 for p in people if p["cat"] == k)))
    print("完整性檢查：各類合計 %d %s 有效名單 %d；有效＋排除＋無法分區 %d %s 報表總人數 %d" % (
        cat_sum, "＝" if cat_sum == n_valid else "≠", n_valid,
        n_valid + n_ex + unassigned, "＝" if n_valid + n_ex + unassigned == res["total_rows"] else "≠", res["total_rows"]))
    if res["warn_areas"]:
        print("登載示警：" + "、".join(res["warn_areas"]))
    if res["issues"]:
        print("資料待確認 %d 筆" % len(res["issues"]))
    for g, _, sheet in CROSS:
        yp = [p for p in people if p["group"] == g]
        print("%s跨區：有效%d人，列入關心%d人（不重複計入全區）" % (sheet, len(yp), sum(1 for p in cross_txt[g] if p["cat"] in CATS)))
    if res["aux"]:
        listed = [p for p in people if p["cat"] in CATS and p["area"] not in res["warn_areas"]]
        print("標＊（近一季小排、家聚會都沒有紀錄）：%d人／列入名單%d人" % (sum(1 for p in listed if p["star"]), len(listed)))
        for c in CATS:
            print("　第%d順位：標＊%d／%d" % (c, sum(1 for p in listed if p["cat"] == c and p["star"]), sum(1 for p in listed if p["cat"] == c)))
        for g, _, sheet in CROSS:
            print("%s標＊：%d人" % (sheet, sum(1 for p in cross_txt[g] if p["cat"] in CATS and p["star"])))
        print("主日沒來但有小排家聚會（內部名單，Excel 與文字名單最下面，不製作卡片）：%d人" % len(res["excl_contact"]))
    print("接近穩定（內部參考，只放 Excel 與文字名單最下面）：%d人" % len(near_people(res)))


if __name__ == "__main__":
    main()
