# -*- coding: utf-8 -*-
"""
近一年出席分析｜規則 V1.1 分析程式

用法：
    只有主日：
    python3 annual_attendance.py 主日報表.xlsx 輸出資料夾
    主日＋小排＋家聚會：
    python3 annual_attendance.py 主日報表.xlsx 輸出資料夾 --group 小排報表.xlsx --home 家聚會報表.xlsx

說明：
- 週次判定沿用 V7；全區出席為0的週次一律視為未點名、不納入。
- 近一年＝最近一次主日往前364天內的已舉行主日；最近48次分成4段，每段12次。「年初」＝第1段，「年底」＝最後一段。
- 每段分級：穩定9～12次、普通6～8次、不常來1～5次、沒來0次；「常來」＝穩定或普通。
- 會所與各區人數不含訪客（臨時點名）與兒童（兒童專項，及各區群組為學齡前、國小者，含已得救小高）；兒童另列。
- 判斷「留住」「已穩定連結」時同時看近一個月（近4週）：近4週都沒來者另列，不算留住或已穩定連結。
- 「近12週主日都沒來，但小排或家聚會有來」不限近一年主日次數（2026-11-13 修訂）。
- 只呈現人數、比例與「待了解」訊號，不判定原因。
"""
import re
import os
import sys
import argparse
import datetime as dt
from collections import OrderedDict, Counter

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

AREAS = ["1小區", "2小區", "3小區", "4小區", "5小區", "6小區", "7小區", "青年區", "兒童專項"]
ADULT_AREAS = AREAS[:-1]
KID_GROUPS = ("學齡前", "國小")
EXCLUDE_KEYWORDS = ["來訪", "訪客", "臨時", "福音朋友"]
CROSS = [("青職", "青職（跨區）"), ("大學", "大專（跨區）")]
ALL = "全會所"
KIDS = "兒童"
SCOPE_NOTE = ("會所與各區人數不含訪客（姓名或欄位含來訪、訪客、臨時、福音朋友的臨時點名）"
              "與兒童（兒童專項，及各區群組為學齡前、國小者，含已得救小高）；兒童、訪客另列。")
SEG = 12
YEAR_DAYS = 364
TIERS = ["穩定", "普通", "不常來", "沒來"]
TIER_CH = {"穩定": "穩", "普通": "普", "不常來": "少", "沒來": "無"}
OFTEN = ("穩定", "普通")
TIER_NOTE = "穩定＝12週來9次以上、普通＝6～8次、不常來＝1～5次、沒來＝0次；常來＝穩定或普通"
TYPES = ["留住", "常來但近一個月沒來", "下滑", "淡出", "回升", "本區間第二季以後才開始出現", "一直不常來"]
TYPE_NOTE = OrderedDict([
    ("留住", "年初常來，年底仍常來，近一個月有來"),
    ("常來但近一個月沒來", "年底仍算常來（12週6次以上），但近一個月（4週）都沒來"),
    ("下滑", "年初常來，年底變不常來（1～5次）"),
    ("淡出", "年初常來，年底都沒來"),
    ("回升", "年初不常來或沒來，年底變常來，近一個月有來"),
    ("本區間第二季以後才開始出現", "本區間第一季（第1段）及更早週次都沒有出席，第二季以後才第一次出席"),
    ("一直不常來", "年初、年底都是不常來或沒來"),
])
NEW_STATUS = ["已穩定連結", "曾常來但近一個月沒來", "有再來但未穩定", "只來1～2次", "觀察中"]
NEW_NOTE = ("出現後12次主日來6次以上且近一個月有來＝已穩定連結；來6次以上但近一個月都沒來＝曾常來但近一個月沒來；"
            "3～5次＝有再來但未穩定；1～2次＝只來1～2次（可能多為來訪性質）；第一次出席後未滿12週＝觀察中。")
DROP_MIN = 3          # 人數減少的區：列出年底比年初少來3次以上者
SIG_DROP = 0.10
SIG_RET_GAP = 0.10
SIG_FADE = 0.15
SIG_MIN_BASE = 10
LIMITS = [
    "報表是匯出當下的名單；年中已從系統刪除的人不在報表內，年初人數與流失可能被低估、留存率可能偏高。",
    "第二季以後遷入者會列在「本區間第二季以後才開始出現」；第一季遷入者照一般類型分類。",
    "點名漏登會被當成缺席。",
    "數字只用來提出值得了解的問題，不判定原因，不據以評價區、負責弟兄或個人。",
]


def fail(msg):
    print("【停止】" + msg)
    sys.exit(2)


def tier(n):
    return "穩定" if n >= 9 else ("普通" if n >= 6 else ("不常來" if n >= 1 else "沒來"))


def read_report(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["出席"] if "出席" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    ref_date = header_idx = None
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
        fail("%s 找不到「製表日期」。" % path)
    if header_idx is None:
        fail("%s 找不到含「姓名」的標題列。" % path)
    header = [str(x).strip() if x is not None else "" for x in rows[header_idx]]
    col = {n: header.index(n) for n in ["姓名", "大區", "小區", "群組"] if n in header}
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


def pick_weeks(ref_date, date_cols, data):
    notes = []
    by_mon = OrderedDict()
    for j, mon in date_cols:
        by_mon.setdefault(mon, []).append(j)
    held = []
    for mon, js in by_mon.items():
        sunday = mon + dt.timedelta(days=6)
        filled = [j for j in js if sum(int(r[j] or 0) for r in data) > 0]
        if len(filled) > 1:
            fail("%s 主日有%d個重複週次欄且都有出席資料，請確認點名系統後重新匯出。" % (sunday, len(js)))
        if sunday >= ref_date:
            notes.append("%s 主日尚未舉行（製表日期 %s），不納入。" % (sunday, ref_date))
            continue
        if not filled:
            notes.append("%s 主日全區出席為0，視為未點名，不納入。" % sunday)
            continue
        if len(js) > 1:
            notes.append("%s 主日報表有重複週次欄（%d欄），採用有出席資料的那一欄。" % (sunday, len(js)))
        held.append((filled[0], sunday))
    if not held:
        fail("報表中沒有已舉行的主日。")
    last = held[-1][1]
    start = last - dt.timedelta(days=YEAR_DAYS)
    year = [x for x in held if x[1] > start]
    earlier = [x for x in held if x[1] <= start]
    if len(year) < 2 * SEG:
        fail("近一年內已舉行主日只有%d次，不足兩段（24次），無法做近一年分析。" % len(year))
    if held[0][1] > start + dt.timedelta(days=7):
        notes.append("報表期間不足一年（最早主日 %s），近一年分析以現有週次進行。" % held[0][1])
    return year, earlier, notes


def area_of(r, col):
    big = (r[col["大區"]] or "").strip()
    small = (r[col["小區"]] or "").strip() if r[col["小區"]] else ""
    if big == "會所青年區":
        return "青年區", None
    if big == "兒童專項":
        return "兒童專項", ("大區為兒童專項、小區填青年區，依規則歸兒童專項" if small == "青年區" else None)
    if small in AREAS:
        return small, None
    return None, "無法判定分區（大區：%s，小區：%s）" % (big or "空白", small or "空白")


def row_key(r, col):
    return (str(r[col["姓名"]]).strip(), str(r[col["大區"]] or "").strip(), str(r[col["小區"]] or "").strip())


def load_aux(path, kind, ref_date, mons, sunday_total):
    aref, acol, adates, adata = read_report(path)
    if aref != ref_date:
        fail("%s報表製表日期 %s 與主日報表 %s 不同，請匯出同一天的報表。" % (kind, aref, ref_date))
    have = set(d for _, d in adates)
    missing = [m for m in mons if m not in have]
    if missing:
        fail("%s報表缺少近一年週次：%s" % (kind, "、".join(m.isoformat() for m in missing)))
    recs, cnt = {}, Counter()
    for r in adata:
        k = row_key(r, acol)
        cnt[k] += 1
        recs[k + (cnt[k],)] = [max(int(r[j] or 0) for j, d in adates if d == m) for m in mons]
    total = sum(sum(v) for v in recs.values())
    if total >= sunday_total:
        fail("%s報表近一年出席總數（%d）不少於主日報表（%d），檔案可能放錯。" % (kind, total, sunday_total))
    return recs, cnt


def classify_type(tiers, r4, newcomer):
    t0, tl = tiers[0], tiers[-1]
    if newcomer:
        return "本區間第二季以後才開始出現"
    if tl in OFTEN and r4 == 0:
        return "常來但近一個月沒來"
    if t0 in OFTEN:
        return {"穩定": "留住", "普通": "留住", "不常來": "下滑", "沒來": "淡出"}[tl]
    return "回升" if tl in OFTEN else "一直不常來"


def new_status(n12, r4, observing):
    if observing:
        return "觀察中"
    if n12 >= 6:
        return "已穩定連結" if r4 > 0 else "曾常來但近一個月沒來"
    return "有再來但未穩定" if n12 >= 3 else "只來1～2次"


def build(path, group=None, home=None):
    ref_date, col, date_cols, data = read_report(path)
    year, earlier, notes = pick_weeks(ref_date, date_cols, data)
    W = len(year)
    nseg = min(4, W // SEG)
    seg_start = W - nseg * SEG
    sundays = [s for _, s in year]
    pre_cols = [j for j, _ in earlier]

    people, issues = [], []
    zero_people = []  # 近一年主日0次者（三份報表模式比對小排、家聚會用）
    placeholder, zero_year = Counter(), Counter()
    trend_place = [0] * W
    scnt = Counter()
    for r in data:
        rk = row_key(r, col)
        scnt[rk] += 1
        key = rk + (scnt[rk],)
        name = rk[0]
        texts = " ".join(str(r[col[c]] or "") for c in col)
        a, issue = area_of(r, col)
        if any(kw in texts for kw in EXCLUDE_KEYWORDS):
            placeholder[a or "未分區"] += 1
            for i, (j, _) in enumerate(year):
                trend_place[i] += int(r[j] or 0)
            continue
        seq = [1 if int(r[j] or 0) else 0 for j, _ in year]
        if sum(seq) == 0:
            zero_year[a or "未分區"] += 1
            if a is not None:
                zero_people.append(dict(name=name, area=a, key=key, group=(r[col["群組"]] if "群組" in col else "") or ""))
            continue
        if issue:
            issues.append("%s：%s" % (name, issue))
        if a is None:
            continue
        grp = (r[col["群組"]] if "群組" in col else "") or ""
        segs = [sum(seq[seg_start + i * SEG: seg_start + (i + 1) * SEG]) for i in range(nseg)]
        tiers = [tier(n) for n in segs]
        first = seq.index(1)
        newcomer = first >= seg_start + SEG and not any(int(r[j] or 0) for j in pre_cols)
        r4 = sum(seq[-4:])
        p = dict(name=name, area=a, group=grp, kid=(a == "兒童專項" or grp in KID_GROUPS), key=key,
                 seq=seq, segs=segs, tiers=tiers, r4=r4, type=classify_type(tiers, r4, newcomer), total=sum(seq),
                 last=max(sundays[i] for i, x in enumerate(seq) if x), first=first, newcomer=newcomer)
        if newcomer:
            obs = first + SEG > W
            p["new_n"] = sum(seq[first:first + SEG])
            p["new_status"] = new_status(p["new_n"], r4, obs)
        people.append(p)

    aux = bool(group or home)
    no_sunday = []
    if aux:
        if not (group and home):
            fail("小排、家聚會報表需兩份一起提供（--group 與 --home）。")
        mons = [s - dt.timedelta(days=6) for s in sundays]
        sunday_total = sum(int(r[j] or 0) for r in data for j, _ in year)
        g_recs, g_cnt = load_aux(group, "小排", ref_date, mons, sunday_total)
        h_recs, h_cnt = load_aux(home, "家聚會", ref_date, mons, sunday_total)
        for p in people:
            rk = p["key"][:3]
            if g_cnt[rk] == 0 and h_cnt[rk] == 0:
                p["gh"] = [0] * W
                issues.append("%s（%s）：小排、家聚會報表中找不到此人，視為無紀錄" % (p["name"], p["area"]))
            elif g_cnt[rk] != scnt[rk] or h_cnt[rk] != scnt[rk]:
                p["gh"] = None
                issues.append("%s（%s）：同區同名筆數在三份報表中不一致，無法比對小排、家聚會" % (p["name"], p["area"]))
            else:
                p["gh"] = [1 if (x or y) else 0 for x, y in zip(g_recs[p["key"]], h_recs[p["key"]])]
            if p["gh"] is not None:
                p["gh_segs"] = [sum(p["gh"][seg_start + i * SEG: seg_start + (i + 1) * SEG]) for i in range(nseg)]
                if p["newcomer"] and p["new_status"] != "觀察中":
                    p["new_gh"] = 1 if any(p["gh"][p["first"]:p["first"] + SEG]) else 0
        # 近12週主日都沒來，但小排或家聚會有來（2026-11-13 修訂：不限近一年主日次數，
        # 最近一季主日0次，且最近一季小排或家聚會有紀錄；同區同名筆數不一致者無法比對，不列）
        last_seg = slice(seg_start + (nseg - 1) * SEG, W)
        cands = [dict(e, sy=0) for e in zero_people]
        cands += [dict(name=p["name"], area=p["area"], key=p["key"], group=p["group"], sy=p["total"])
                  for p in people if p["segs"][-1] == 0]
        for e in cands:
            rk = e["key"][:3]
            if g_cnt[rk] != scnt[rk] or h_cnt[rk] != scnt[rk]:
                continue
            g, h = g_recs[e["key"]], h_recs[e["key"]]
            e.update(gy=sum(1 for x in g if x), hy=sum(1 for x in h if x),
                     gq=sum(1 for x in g[last_seg] if x), hq=sum(1 for x in h[last_seg] if x))
            if e["gq"] or e["hq"]:
                no_sunday.append(e)
        order = {a: i for i, a in enumerate(AREAS)}
        no_sunday.sort(key=lambda e: (order[e["area"]], -(e["gq"] + e["hq"]), -(e["gy"] + e["hy"]), e["name"]))

    return dict(ref_date=ref_date, sundays=sundays, notes=notes, nseg=nseg, seg_start=seg_start, no_sunday=no_sunday,
                people=people, issues=issues, placeholder=placeholder, zero_year=zero_year,
                trend_place=trend_place, total_rows=len(data), aux=aux)


# ---------- 統計 ----------

def units(res):
    ps = res["people"]
    out = OrderedDict()
    out[ALL] = [p for p in ps if not p["kid"]]
    for a in ADULT_AREAS:
        out[a] = [p for p in ps if p["area"] == a and not p["kid"]]
    out[KIDS] = [p for p in ps if p["kid"]]
    for g, title in CROSS:
        out[title] = [p for p in ps if p["group"] == g]
    return out


def movement(ps):
    m = OrderedDict((t0, OrderedDict((t1, 0) for t1 in TIERS)) for t0 in TIERS)
    for p in ps:
        m[p["tiers"][0]][p["tiers"][-1]] += 1
    tc = Counter(p["type"] for p in ps)
    base = sum(m[t0][t1] for t0 in OFTEN for t1 in TIERS)
    kept = sum(m[t0][t1] for t0 in OFTEN for t1 in OFTEN)
    kept_away = sum(1 for p in ps if p["tiers"][0] in OFTEN and p["tiers"][-1] in OFTEN and p["r4"] == 0)
    return dict(m=m, base=base, kept=kept, kept_away=kept_away, ret=(kept / base if base else None),
                down=sum(m[t0]["不常來"] for t0 in OFTEN), fade=sum(m[t0]["沒來"] for t0 in OFTEN),
                up=sum(1 for p in ps if not p["newcomer"] and p["tiers"][0] not in OFTEN and p["tiers"][-1] in OFTEN), tc=tc)


def seg_stats(res, ps):
    nseg = res["nseg"]
    return ([sum(p["segs"][i] for p in ps) / SEG for i in range(nseg)],
            [sum(1 for p in ps if p["segs"][i] > 0) for i in range(nseg)],
            [sum(1 for p in ps if p["segs"][i] >= 9) for i in range(nseg)])


def contrib(ps):
    """年底比年初每週平均增減，依軌跡類型拆開（各類加總＝總增減）。"""
    c = OrderedDict((t, 0) for t in TYPES)
    for p in ps:
        c[p["type"]] += p["segs"][-1] - p["segs"][0]
    return OrderedDict((t, v / SEG) for t, v in c.items())


def main_drop(c):
    t = min(c, key=lambda k: c[k])
    return (t, c[t]) if c[t] < 0 else (None, 0)


def weekly(res):
    U = units(res)
    W = len(res["sundays"])
    rows = []
    for i in range(W):
        per = OrderedDict((a, sum(p["seq"][i] for p in U[a])) for a in ADULT_AREAS)
        rows.append((res["sundays"][i], sum(p["seq"][i] for p in U[ALL]), per,
                     sum(p["seq"][i] for p in U[KIDS]), res["trend_place"][i]))
    ma = [sum(r[1] for r in rows[i - 3:i + 1]) / 4 if i >= 3 else None for i in range(W)]
    return rows, ma


COMPARE_SHEET = "前後季對照"


def month_span(res, i):
    """第i段的實際月份，例如「12～3月」（跨年不寫年份）。"""
    s = res["sundays"][res["seg_start"] + i * SEG]
    e = res["sundays"][res["seg_start"] + (i + 1) * SEG - 1]
    return "%d～%d月" % (s.month, e.month)


def by_month(res, text):
    """把「年初」「年底」換成最早一季、最近一季的實際月份。"""
    return text.replace("年初", month_span(res, 0)).replace("年底", month_span(res, res["nseg"] - 1))


def seg_label(res, i):
    s = res["sundays"][res["seg_start"] + i * SEG]
    e = res["sundays"][res["seg_start"] + (i + 1) * SEG - 1]
    return "第%d段（%s～%s）" % (i + 1, s.strftime("%Y/%m/%d"), e.strftime("%Y/%m/%d"))


def week_seg(res, i):
    k = i - res["seg_start"]
    return "分段前" if k < 0 else "第%d段" % (k // SEG + 1)


def decline_areas(res):
    U = units(res)
    out = []
    for a in ADULT_AREAS:
        avg, _, _ = seg_stats(res, U[a])
        if avg[0] > 0 and (avg[-1] - avg[0]) / avg[0] <= -SIG_DROP:
            out.append(a)
    return out


def growth_areas(res):
    """年底比年初增加5%以上的單位（各區與兒童）。"""
    U = units(res)
    out = []
    for u in list(ADULT_AREAS) + [KIDS]:
        avg, _, _ = seg_stats(res, U[u])
        if avg[0] and (avg[-1] - avg[0]) / avg[0] >= HL_GROW:
            out.append(u)
    return out


def split_change(ps, up):
    """把年底比年初的增減拆成「明顯變化的人（多來或少來3次以上）」與「其餘（許多人各一兩次）」。"""
    d = lambda p: p["segs"][-1] - p["segs"][0]
    big = [p for p in ps if (d(p) >= DROP_MIN if up else -d(p) >= DROP_MIN)]
    total = sum(d(p) for p in ps) / SEG
    big_sum = sum(d(p) for p in big) / SEG
    often = sum(1 for p in big if (p["tiers"][-1] if up else p["tiers"][0]) in OFTEN)
    return dict(big=big, n=len(big), big_sum=big_sum, rest=total - big_sum, often=often)


def split_text(ps, up):
    x = split_change(ps, up)
    if up:
        head = "多來%d次以上的%d人合計%+.1f人（其中年底已常來%d人），其餘%+.1f人" % (DROP_MIN, x["n"], x["big_sum"], x["often"], x["rest"])
    else:
        head = "少來%d次以上的%d人合計%+.1f人（其中年初常來%d人），其餘%+.1f人" % (DROP_MIN, x["n"], x["big_sum"], x["often"], x["rest"])
    return x, head


def split_judge(ps, up):
    x, head = split_text(ps, up)
    concentrated = abs(x["big_sum"]) >= abs(x["rest"])
    if up:
        if concentrated:
            tail = "代表有人變穩了" if x["often"] * 2 >= x["n"] else "代表有人明顯多來了，但多數還不算常來"
            return "增加集中在少數人：%s。%s。" % (head, tail)
        return "增加分散：%s來自許多人各多來一兩次，代表輪流來的人變多了。" % head
    if concentrated:
        return "減少集中在少數人：%s。建議負責弟兄逐一了解這些人。" % head
    return "減少分散：%s來自許多人各少來一兩次，代表輪流來的人整體來得少了；建議從整體連結來看（例如聯絡頻率、小排、接送或聚會時段），不必只找個人。" % head


def signals(res):
    U = units(res)
    allmv = movement(U[ALL])
    out = []
    for a in ADULT_AREAS:
        ps = U[a]
        avg, _, _ = seg_stats(res, ps)
        if a in decline_areas(res):
            t, v = main_drop(contrib(ps))
            out.append("%s：年底平均每週%.1f人，比年初的%.1f人減少%.0f%%；減少最多來自「%s」（%+.1f人）" % (
                a, avg[-1], avg[0], (avg[0] - avg[-1]) / avg[0] * 100, t, v))
        mv = movement(ps)
        if mv["base"] >= SIG_MIN_BASE and mv["ret"] <= allmv["ret"] - SIG_RET_GAP:
            out.append("%s：年初常來的人年底仍常來的比例%.0f%%（%d／%d），比全會所%.0f%%低%.0f個百分點" % (
                a, mv["ret"] * 100, mv["kept"], mv["base"], allmv["ret"] * 100, (allmv["ret"] - mv["ret"]) * 100))
        if mv["base"] >= SIG_MIN_BASE and mv["fade"] / mv["base"] >= SIG_FADE:
            out.append("%s：年初常來的%d人中，有%d人（%.0f%%）年底都沒來" % (a, mv["base"], mv["fade"], mv["fade"] / mv["base"] * 100))
    if res["aux"]:
        fade = [p for p in U[ALL] if p["type"] == "淡出" and p.get("gh_segs") is not None]
        out.append("全會所：年底主日都沒來的%d人中，有%d人年底小排、家聚會也都沒有紀錄" % (len(fade), sum(1 for p in fade if not p["gh_segs"][-1])))
    return out


def pct(x):
    return "—" if x is None else "%.0f%%" % (x * 100)


def summary_text(res):
    U = units(res)
    s = res["sundays"]
    nseg = res["nseg"]
    L = ["會所主日近一年出席觀察（分隔線以上不含姓名，可用於同工交通；分隔線以下為內部名單）",
         "製表日期：%s" % res["ref_date"],
         "近一年期間：%s～%s，已舉行主日%d次" % (s[0].strftime("%Y/%m/%d"), s[-1].strftime("%Y/%m/%d"), len(s)),
         "分段：" + "、".join(seg_label(res, i) for i in range(nseg)),
         "等級：" + TIER_NOTE,
         "人數範圍：" + SCOPE_NOTE,
         "分析模式：" + ("主日＋小排＋家聚會" if res["aux"] else "只有主日"), ""]
    rows, _ = weekly(res)
    tots = [r[1] for r in rows]
    avg, att, stab = seg_stats(res, U[ALL])
    L.append("【一、全會所人數走向】")
    L.append("全年平均每週%.1f人；最高%d人（%s），最低%d人（%s）" % (
        sum(tots) / len(tots), max(tots), rows[tots.index(max(tots))][0].strftime("%m/%d"),
        min(tots), rows[tots.index(min(tots))][0].strftime("%m/%d")))
    L.append("各段平均每週：" + "、".join("第%d段%.1f人" % (i + 1, v) for i, v in enumerate(avg)))
    L.append("年底比年初：%+.1f人（%+.0f%%）" % (avg[-1] - avg[0], (avg[-1] - avg[0]) / avg[0] * 100 if avg[0] else 0))
    L.append("各段穩定出席人數：" + "、".join("第%d段%d人" % (i + 1, stab[i]) for i in range(nseg)))
    L.append("")
    mv = movement(U[ALL])
    L.append("【二、年初常來的人，年底怎麼樣】")
    L.append("年初常來共%d人，年底：" % mv["base"])
    L.append("・仍常來%d人（%s），其中近一個月都沒來%d人" % (mv["kept"], pct(mv["ret"]), mv["kept_away"]))
    L.append("・變不常來（下滑）%d人" % mv["down"])
    L.append("・都沒來（淡出）%d人" % mv["fade"])
    L.append("年初不常來或沒來、年底變常來（回升，不含本區間第二季以後才開始出現）：%d人" % mv["up"])
    L.append("每人一種類型：" + "、".join("%s%d人" % (t, mv["tc"][t]) for t in TYPES))
    L.append("")
    L.append("【三、各區比較】")
    L.append("（平均每週人數年初→年底；年初常來者年底仍常來的比例；年底都沒來人數）")
    for u in list(ADULT_AREAS) + [KIDS] + [t for _, t in CROSS]:
        a_avg, _, _ = seg_stats(res, U[u])
        m = movement(U[u])
        L.append("%s：%.1f→%.1f人；仍常來%s（%d／%d）；都沒來%d人" % (u, a_avg[0], a_avg[-1], pct(m["ret"]), m["kept"], m["base"], m["fade"]))
    L.append("")
    L.append("【四、各區增減來自哪些人】（每週平均人數，年底減年初；各類加總＝總增減；青職、大專為跨區，不另計入全會所）")
    for a in list(ADULT_AREAS) + [KIDS] + [t for _, t in CROSS]:
        c = contrib(U[a])
        tot = sum(c.values())
        parts = "、".join("%s%+.1f" % (t, v) for t, v in c.items() if abs(v) >= 0.05)
        L.append("%s（%+.1f）：%s" % (a, tot, parts or "無變化"))
    L.append("")
    L.append("【五、本區間第二季以後才開始出現】")
    nw = [p for p in U[ALL] if p["newcomer"]]
    nc = Counter(p["new_status"] for p in nw)
    L.append("共%d人：" % len(nw) + "、".join("%s%d人" % (st, nc[st]) for st in NEW_STATUS))
    L.append("各區（" + "／".join(NEW_STATUS) + "）：" + "、".join(
        "%s %s" % (a, "／".join(str(sum(1 for p in U[a] if p["newcomer"] and p["new_status"] == st)) for st in NEW_STATUS)) for a in ADULT_AREAS))
    L.append("")
    if res["aux"]:
        L.append("【六、主日與小排、家聚會】")
        dd = [p for p in U[ALL] if p["type"] in ("下滑", "淡出") and p.get("gh_segs") is not None]
        L.append("主日下滑或淡出的%d人中，年底小排或家聚會仍有紀錄%d人" % (len(dd), sum(1 for p in dd if p["gh_segs"][-1])))
        ns = res["no_sunday"]
        L.append("近12週主日都沒來，但年底小排或家聚會有紀錄：%d人，其中年底來%d週以上%d人（名單見內部名單第四段）" % (
            len(ns), HL_INVITE, sum(1 for e in ns if e["gq"] + e["hq"] >= HL_INVITE)))
        j = [p for p in nw if "new_gh" in p]
        for flag, lab in ((1, "有"), (0, "沒有")):
            grp = [p for p in j if p["new_gh"] == flag]
            ok = sum(1 for p in grp if p["new_status"] == "已穩定連結")
            L.append("本區間第二季以後才開始出現、出現後12週%s小排或家聚會紀錄：%d人，其中已穩定連結%d人（%s）" % (lab, len(grp), ok, pct(ok / len(grp) if grp else None)))
        L.append("")
    L.append("【待了解訊號】（只提出值得了解的地方，不代表原因或評價）")
    sig = signals(res)
    L += ["・" + x for x in sig] if sig else ["無"]
    L.append("")
    L.append("【資料限制】")
    L += ["・" + x for x in LIMITS]
    return "\n".join(L) + "\n"


INTERNAL_SEP = "\n" + "＝" * 20 + "\n\n"
SPLIT_NOTE = ("說明：每區的增減拆成兩部分：「多來（少來）3次以上的人合計」與「其餘」（許多人各多來或少來一兩次的總和）。"
              "前者大，表示變化集中在少數人；其餘大，表示是輪流來的人整體變多或變少。")
NO_SUNDAY_NOTE = ("近12週（年底）主日都沒來，但年底小排或家聚會有紀錄（不限近一年主日次數）。"
                  "仍有連結，可邀請來主日。同區依年底小排、家聚會來得多者排前面；近一年主日次數多者，可能原本常來、最近一季才沒來。")


def no_sunday_line(e):
    return "　%s（%s｜近一年主日%d次｜近一年小排%d週、家聚會%d週｜年底小排%d週、家聚會%d週）" % (
        e["name"], e["group"] or "無群組", e["sy"], e["gy"], e["hy"], e["gq"], e["hq"])
INTERNAL_TITLE = "近一年內部名單｜內部參考，不對外轉發"


def internal_text(res):
    """文字摘要最下面的內部名單（分隔線以上可轉發，以下含姓名）。"""
    U = units(res)
    order = {a: i for i, a in enumerate(AREAS)}
    tyo = {t: i for i, t in enumerate(TYPES)}
    L = [INTERNAL_TITLE, "各段為12週出席次數，依序為%s；近4週＝近一個月。" % "→".join(month_span(res, i) for i in range(res["nseg"])), ""]

    def gh(p):
        if not res["aux"]:
            return ""
        return "｜年底小排家聚會%s" % ("無法比對" if p.get("gh_segs") is None else "%d週" % p["gh_segs"][-1])

    L.append("【一、下滑、淡出、常來但近一個月沒來】")
    L.append("；".join("%s＝%s" % (t, TYPE_NOTE[t]) for t in ("下滑", "淡出", "常來但近一個月沒來")))
    lst = [p for p in U[ALL] + U[KIDS] if p["type"] in ("下滑", "淡出", "常來但近一個月沒來")]
    groups = [(a, [p for p in lst if not p["kid"] and p["area"] == a]) for a in ADULT_AREAS] + [(KIDS, [p for p in lst if p["kid"]])]
    for a, ps in groups:
        if not ps:
            continue
        L.append("")
        L.append("〔%s〕（%d人）" % (a, len(ps)))
        for p in sorted(ps, key=lambda p: (tyo[p["type"]], p["segs"][-1] - p["segs"][0], p["name"])):
            L.append("　%s（%s｜%s｜%s｜近4週%d次%s）" % (p["name"], p["group"] or "無群組", p["type"],
                                                   "→".join(str(x) for x in p["segs"]), p["r4"], gh(p)))
    if not lst:
        L.append("無")
    L.append("")
    L.append("【二、人數減少10%%以上的區：年底比年初少來%d次以上的人】" % DROP_MIN)
    L.append(SPLIT_NOTE)
    dec = decline_areas(res)
    if not dec:
        L.append("本次沒有人數減少10%以上的區。")
    for a in dec:
        avg, _, _ = seg_stats(res, U[a])
        t, v = main_drop(contrib(U[a]))
        ps = sorted([p for p in U[a] if p["segs"][0] - p["segs"][-1] >= DROP_MIN], key=lambda p: (p["segs"][-1] - p["segs"][0], p["name"]))
        L.append("")
        L.append("〔%s〕年初平均%.1f人→年底平均%.1f人；減少最多來自「%s」（%+.1f人）；%s" % (a, avg[0], avg[-1], t, v, split_text(U[a], False)[1]))
        for p in ps:
            L.append("　%s（%s｜%s｜年初來%d次→年底來%d次，少來%d次｜近4週%d次%s）" % (
                p["name"], p["group"] or "無群組", p["type"], p["segs"][0], p["segs"][-1], p["segs"][0] - p["segs"][-1], p["r4"], gh(p)))
    L.append("")
    L.append("【三、人數增加5%%以上的區：年底比年初多來%d次以上的人】" % DROP_MIN)
    L.append(SPLIT_NOTE)
    grow = growth_areas(res)
    if not grow:
        L.append("本次沒有人數增加5%以上的區。")
    for a in grow:
        avg, _, _ = seg_stats(res, U[a])
        c = contrib(U[a])
        t = max(c, key=lambda k: c[k])
        ps = sorted(split_change(U[a], True)["big"], key=lambda p: (p["segs"][0] - p["segs"][-1], p["name"]))
        L.append("")
        L.append("〔%s〕年初平均%.1f人→年底平均%.1f人；增加最多來自「%s」（%+.1f人）；%s" % (a, avg[0], avg[-1], t, c[t], split_text(U[a], True)[1]))
        for p in ps:
            L.append("　%s（%s｜%s｜年初來%d次→年底來%d次，多來%d次｜近4週%d次%s）" % (
                p["name"], p["group"] or "無群組", p["type"], p["segs"][0], p["segs"][-1], p["segs"][-1] - p["segs"][0], p["r4"], gh(p)))
    if res["aux"]:
        L.append("")
        L.append("【四、近12週主日都沒來，但小排或家聚會有來】")
        L.append(NO_SUNDAY_NOTE)
        ns = res["no_sunday"]
        L.append("共%d人" % len(ns))
        for a in AREAS:
            lst = [e for e in ns if e["area"] == a]
            if not lst:
                continue
            L.append("")
            L.append("〔%s〕（%d人）" % (a, len(lst)))
            for e in lst:
                L.append(no_sunday_line(e))
    return "\n".join(L) + "\n"



# ---------- 重點摘要（2026-11-13 新增） ----------
HL_TITLE = "近一年出席重點摘要｜內部參考，不對外轉發"
HL_FLAT = 0.03     # 增減在3%以內視為大致持平
HL_GROW = 0.05     # 增加5%以上列為人數上升的區
HL_INVITE = 2      # 主日都沒來、年底小排家聚會來2週以上，列為可邀回主日
SOURCE_ADVICE = OrderedDict([
    ("下滑", "減少最多來自原本常來、年底變得不常來的人（下滑）。"),
    ("淡出", "減少最多來自原本常來、年底都沒來的人（淡出）。"),
    ("常來但近一個月沒來", "減少最多來自仍算常來、但近一個月都沒來的人。"),
    ("一直不常來", "減少最多來自本來就不常來的人來得更少（一直不常來）。"),
    ("本區間第二季以後才開始出現", "減少最多來自第二季以後新出現、之後沒有留下來的人。"),
    ("留住", "減少最多來自固定來的人都來得少一些（留住）。"),
    ("回升", "減少最多來自回升的人又來得較少。"),
])


def hl_lists(res):
    """重點摘要第五、六點的名單（回測程式獨立重算核對）。"""
    U = units(res)
    order = {a: i for i, a in enumerate(AREAS)}
    ps = U[ALL] + U[KIDS]
    key = lambda p: (order[p["area"]], p["name"])
    fade = sorted([p for p in ps if p["type"] == "淡出"], key=key)
    away = sorted([p for p in ps if p["type"] == "常來但近一個月沒來"], key=key)
    down0 = sorted([p for p in ps if p["type"] == "下滑" and p["r4"] == 0], key=key)
    new_away = sorted([p for p in ps if p["newcomer"] and p["new_status"] == "曾常來但近一個月沒來"], key=key)
    invite = [e for e in res["no_sunday"] if e["gq"] + e["hq"] >= HL_INVITE] if res["aux"] else []
    return fade, away, down0, new_away, invite


def highlights_text(res):
    U = units(res)
    allmv = movement(U[ALL])
    avg, _, stab = seg_stats(res, U[ALL])
    chg = (avg[-1] - avg[0]) / avg[0] if avg[0] else 0
    L = [HL_TITLE, "製表日期：%s" % res["ref_date"],
         "比較：年初（最早一季）與年底（最近一季）；近一個月＝最近4次主日",
         "說明：依固定規則從分析結果挑出值得先留意的地方，只是提出問題，不代表原因，也不據以評價區或個人。人數不含訪客、兒童。", ""]

    L.append("【一、整體情形】")
    trend = "大致持平" if abs(chg) < HL_FLAT else ("增加" if chg > 0 else "減少")
    L.append("・全會所主日平均每週年初%.1f人→年底%.1f人（%+.0f%%），%s；穩定出席人數年初%d人→年底%d人。" % (
        avg[0], avg[-1], chg * 100, trend, stab[0], stab[-1]))
    L.append("・年初常來的%d人，年底仍常來%d人（%s）；變不常來%d人、都沒來%d人。" % (
        allmv["base"], allmv["kept"], pct(allmv["ret"]), allmv["down"], allmv["fade"]))
    nw = [p for p in U[ALL] if p["newcomer"]]
    judged = [p for p in nw if p["new_status"] != "觀察中"]
    ok = sum(1 for p in judged if p["new_status"] == "已穩定連結")
    L.append("・本區間第二季以後才開始出現%d人；出現滿12週的%d人中，已穩定連結%d人（%s）。" % (
        len(nw), len(judged), ok, pct(ok / len(judged) if judged else None)))
    L.append("")

    L.append("【二、人數上升的區：誰多來了】")
    good = []
    for u in growth_areas(res):
        a_avg, _, _ = seg_stats(res, U[u])
        if True:
            c = contrib(U[u])
            t = max(c, key=lambda k: c[k])
            good.append("・%s：平均每週年初%.1f人→年底%.1f人（%+.0f%%），增加最多來自「%s」（%+.1f人）。" % (
                u, a_avg[0], a_avg[-1], (a_avg[-1] - a_avg[0]) / a_avg[0] * 100, t, c[t]))
            good.append("　" + split_judge(U[u], True) + "名單見「近一年出席觀察」最下面內部名單【三】（Excel：近一年概況第四段）。")
    L += good or ["・本次沒有人數增加5%以上的區。"]
    if good:
        L.append("　建議了解這些人多來的原因。")
    L.append("・年初不常來、年底變常來（回升）%d人。" % allmv["up"])
    L.append("")

    L.append("【三、人數下降的區：誰少來了】")
    need = []
    for a in decline_areas(res):
        a_avg, _, _ = seg_stats(res, U[a])
        t, v = main_drop(contrib(U[a]))
        need.append("・%s：平均每週年初%.1f人→年底%.1f人（%+.0f%%）。%s" % (
            a, a_avg[0], a_avg[-1], (a_avg[-1] - a_avg[0]) / a_avg[0] * 100, SOURCE_ADVICE[t]))
        need.append("　" + split_judge(U[a], False) + "名單見「近一年出席觀察」最下面內部名單【二】（Excel：近一年概況第三段）。")
    for u in list(ADULT_AREAS) + [t for _, t in CROSS]:
        mv = movement(U[u])
        if mv["base"] >= SIG_MIN_BASE and mv["ret"] <= allmv["ret"] - SIG_RET_GAP:
            need.append("・%s：年初常來的%d人，年底仍常來只有%s（全會所%s）；變不常來%d人、都沒來%d人，建議了解這群人的情況。" % (
                u, mv["base"], pct(mv["ret"]), pct(allmv["ret"]), mv["down"], mv["fade"]))
    L += need or ["・本次沒有人數下降或需要留意的區。"]
    L.append("")

    fade, away, down0, new_away, invite = hl_lists(res)
    L.append("【四、新出現的人】")
    obs = sum(1 for p in nw if p["new_status"] == "觀察中")
    L.append("・已穩定連結比例%s；還在觀察期（出現未滿12週）%d人，建議持續陪伴。" % (pct(ok / len(judged) if judged else None), obs))
    if new_away:
        L.append("・出現後曾經常來、但近一個月都沒來%d人，建議儘快聯絡（名單見第五點）。" % len(new_away))
    L.append("")

    def gh(p):
        if not res["aux"] or p.get("gh_segs") is None:
            return ""
        return "，年底小排家聚會%d週" % p["gh_segs"][-1]

    L.append("【五、建議優先聯絡（內部名單）】")
    L.append("從「近一年出席觀察」內部名單【一】（下滑、淡出、常來但近一個月沒來）與本區間第二季以後才開始出現的人當中，挑出最近一個月都沒來、需要主動聯絡的人。")
    L.append("名單可能包含已遷出或應停用但系統尚未調整的人，發現時請自行略過。")
    n_down = sum(1 for p in U[ALL] + U[KIDS] if p["type"] == "下滑")
    blocks = [("淡出：年初常來，年底都沒來", fade, lambda p: "年初%d次→年底0次%s" % (p["segs"][0], gh(p)),
               "全部列出：原本常來，最近一季完全沒來，最需要了解。"),
              ("常來但近一個月都沒來", away, lambda p: "年底%d次，近一個月0次%s" % (p["segs"][-1], gh(p)),
               "全部列出：斷掉的時間還短，最容易接住。"),
              ("下滑，且近一個月都沒來", down0, lambda p: "年初%d次→年底%d次%s" % (p["segs"][0], p["segs"][-1], gh(p)),
               "下滑共%d人，只列近一個月都沒來的%d人：近一個月還有來的人仍見得到面，可在聚會中關心，不必另外聯絡。" % (n_down, len(down0))),
              ("本區間第二季以後才開始出現，曾常來但近一個月沒來", new_away, lambda p: "出現後12週來%d次%s" % (p["new_n"], gh(p)),
               "從新出現的人中挑出：才剛連上又斷了，最容易流失。")]
    for label, lst, desc, why in blocks:
        L.append("〔%s〕（%d人）%s" % (label, len(lst), why))
        L += ["　%s（%s｜%s）" % (p["name"], p["area"], desc(p)) for p in lst] or ["　無"]
    L.append("")
    if res["aux"]:
        L.append("【六、可以邀回主日（內部名單）】")
        L.append("近12週主日都沒來（不限近一年主日次數），但年底小排或家聚會來了%d週以上；仍有連結，建議透過小排、家聚會邀請來主日。" % HL_INVITE)
        L.append("從「近一年出席觀察」內部名單【四】的%d人中，只列年底來%d週以上的人：最近還常接觸，透過小排、家聚會邀請最自然；其餘只來過1週，接觸較淡。" % (len(res["no_sunday"]), HL_INVITE))
        L.append("〔共%d人〕" % len(invite))
        L += ["　%s（%s｜近一年主日%d次｜年底小排%d週、家聚會%d週）" % (e["name"], e["area"], e["sy"], e["gq"], e["hq"]) for e in invite] or ["　無"]
    return by_month(res, "\n".join(L).rstrip() + "\n")

# ---------- Excel ----------

def write_excel(res, out_path, txt, highlights=""):
    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="BBBBBB")
    hdr_font = Font(bold=True, size=12)
    hdr_fill = PatternFill("solid", fgColor="E7F1FB")
    nseg = res["nseg"]
    U = units(res)
    order = {a: i for i, a in enumerate(AREAS)}
    tyo = {t: i for i, t in enumerate(TYPES)}

    def table(ws, start, headers, rows, widths=None):
        for j, h in enumerate(headers, 1):
            c = ws.cell(row=start, column=j, value=h)
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        for i, row in enumerate(rows, start + 1):
            for j, v in enumerate(row, 1):
                c = ws.cell(row=i, column=j, value=v)
                c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                c.font = Font(size=12)
        if widths:
            for j, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(j)].width = w
        return start + len(rows) + 1

    def title(ws, r, t, size=13, color=None):
        ws.cell(row=r, column=1, value=t).font = Font(bold=True, size=size, color=color)

    def gh_last(p):
        if not res["aux"]:
            return []
        return ["無法比對" if p.get("gh_segs") is None else p["gh_segs"][-1]]

    # 1 近一年概況（上半部同文字摘要；下半部含姓名，內部參考）
    ws = wb.active
    ws.title = "近一年概況"
    r = 0
    for r, t in enumerate(txt.rstrip("\n").split("\n"), 1):
        ws.cell(row=r, column=1, value=t).font = Font(bold=t.startswith("【") or r == 1, size=14 if r == 1 else 12)
    r += 2
    title(ws, r, "＝＝＝ 以下含姓名，只供負責弟兄內部參考，不對外轉發 ＝＝＝", 14, "C00000")
    gh_h = ["年底小排或家聚會週數"] if res["aux"] else []
    seg_h = ["第%d段" % (i + 1) for i in range(nseg)]

    r += 2
    title(ws, r, "一、下滑、淡出、常來但近一個月沒來（依區）")
    ws.cell(row=r + 1, column=1, value="；".join("%s＝%s" % (t, TYPE_NOTE[t]) for t in ("下滑", "淡出", "常來但近一個月沒來")) + "。各段為12週出席次數。")
    lst = sorted([p for p in U[ALL] + U[KIDS] if p["type"] in ("下滑", "淡出", "常來但近一個月沒來")],
                 key=lambda p: (p["kid"], order[p["area"]], tyo[p["type"]], p["segs"][-1] - p["segs"][0], p["name"]))
    rows = [[("兒童" if p["kid"] and p["area"] != "兒童專項" else "") + p["area"], p["name"], p["group"], p["type"]] + p["segs"] + [p["r4"]] + gh_last(p) for p in lst]
    r = table(ws, r + 2, ["所屬區", "姓名", "群組", "類型"] + seg_h + ["近4週"] + gh_h, rows)

    r += 1
    title(ws, r, "二、各區增減來自哪些人（每週平均人數，年底減年初）")
    ws.cell(row=r + 1, column=1, value="把每個人年底比年初多來或少來的次數，依類型加總後除以12，各類加總＝該區總增減。用來看人數減少是少數人下滑、淡出造成，還是不常來的人普遍來得更少。青職、大專為跨區，不另計入全會所。")
    rows = []
    for a in list(ADULT_AREAS) + [KIDS] + [t for _, t in CROSS]:
        c = contrib(U[a])
        avg, _, _ = seg_stats(res, U[a])
        t, v = main_drop(c)
        rows.append([a, round(avg[0], 1), round(avg[-1], 1), round(avg[-1] - avg[0], 1)] + [round(c[x], 1) for x in TYPES] +
                    ["%s（%+.1f）" % (t, v) if t else "—", "是" if a in decline_areas(res) else ""])
    r = table(ws, r + 2, ["區", "年初平均", "年底平均", "增減"] + TYPES + ["減少最多來自", "減少10%以上"], rows,
              [12, 10, 10, 8] + [11] * len(TYPES) + [24, 12])

    r += 1
    dec = decline_areas(res)
    title(ws, r, "三、人數減少10%%以上的區：年底比年初少來%d次以上的人" % DROP_MIN)
    ws.cell(row=r + 1, column=1, value="依少來次數多到少排列（不論類型），可看出減少主要落在哪些人。")
    rows = []
    for a in dec:
        for p in sorted([p for p in U[a] if p["segs"][0] - p["segs"][-1] >= DROP_MIN],
                        key=lambda p: (p["segs"][-1] - p["segs"][0], p["name"])):
            rows.append([a, p["name"], p["group"], p["type"]] + p["segs"] + [p["segs"][-1] - p["segs"][0], p["r4"]] + gh_last(p))
    if not dec:
        ws.cell(row=r + 2, column=1, value="本次沒有人數減少10%以上的區。")
        r += 3
    else:
        r = table(ws, r + 2, ["區", "姓名", "群組", "類型"] + seg_h + ["年底減年初", "近4週"] + gh_h, rows)
    r += 1
    grow = growth_areas(res)
    title(ws, r, "四、人數增加5%%以上的區：年底比年初多來%d次以上的人" % DROP_MIN)
    ws.cell(row=r + 1, column=1, value="依多來次數多到少排列（不論類型），可看出增加是有人變穩了，還是輪流來的人變多了。" + SPLIT_NOTE[3:])
    rows = []
    for a in grow:
        for p in sorted(split_change(U[a], True)["big"], key=lambda p: (p["segs"][0] - p["segs"][-1], p["name"])):
            rows.append([a, p["name"], p["group"], p["type"]] + p["segs"] + [p["segs"][-1] - p["segs"][0], p["r4"]] + gh_last(p))
    if not grow:
        ws.cell(row=r + 2, column=1, value="本次沒有人數增加5%以上的區。")
        r += 3
    else:
        r = table(ws, r + 2, ["區", "姓名", "群組", "類型"] + seg_h + ["年底減年初", "近4週"] + gh_h, rows)
    if res["aux"]:
        r += 1
        title(ws, r, "五、近12週主日都沒來，但小排或家聚會有來")
        ws.cell(row=r + 1, column=1, value=NO_SUNDAY_NOTE)
        rows = [[e["area"], e["name"], e["group"], e["sy"], e["gy"], e["hy"], e["gq"], e["hq"]] for e in res["no_sunday"]]
        if rows:
            r = table(ws, r + 2, ["所屬區", "姓名", "群組", "近一年主日次數", "近一年小排週數", "近一年家聚會週數", "年底小排週數", "年底家聚會週數"], rows)
        else:
            ws.cell(row=r + 2, column=1, value="本次無。")
            r += 3
    ws.column_dimensions["A"].width = 14
    for j, w in zip("BCD", (20, 8, 18)):
        ws.column_dimensions[j].width = w

    # 2 每週趨勢
    ws = wb.create_sheet("每週趨勢")
    ws["A1"] = SCOPE_NOTE
    rows, ma = weekly(res)
    headers = ["主日日期", "段別", "全會所（不含訪客、兒童）", "4週平均"] + ADULT_AREAS + ["兒童（另計）", "訪客等臨時點名（另計）"]
    trows = [[d.isoformat(), week_seg(res, i), tot, (round(ma[i], 1) if ma[i] is not None else None)] + [per[a] for a in ADULT_AREAS] + [kid, pl]
             for i, (d, tot, per, kid, pl) in enumerate(rows)]
    end = table(ws, 2, headers, trows, [12, 8, 14, 9] + [8] * len(ADULT_AREAS) + [10, 12])
    ws.freeze_panes = "C3"
    ch = LineChart()
    ch.title = "全會所每週主日人數（不含訪客、兒童）"
    ch.height, ch.width = 9, 26
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=2, max_row=end - 1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=3, max_row=end - 1))
    ws.add_chart(ch, "P3")
    srt = sorted(range(len(rows)), key=lambda i: rows[i][1])
    r = end + 2
    title(ws, r, "全年最高、最低各5週（原因請自行填寫，程式不推測）")
    hl = [["最高", rows[i][0].isoformat(), rows[i][1], ""] for i in sorted(srt[-5:], key=lambda i: -rows[i][1])]
    hl += [["最低", rows[i][0].isoformat(), rows[i][1], ""] for i in srt[:5]]
    table(ws, r + 1, ["類別", "主日日期", "人數", "原因（請填）"], hl)

    # 3 前後季對照（取代原季度移動表）
    ws = wb.create_sheet(COMPARE_SHEET)
    ws["A1"] = "最早一季＝%s；最近一季＝%s。%s。" % (seg_label(res, 0), seg_label(res, nseg - 1), TIER_NOTE)
    ws["A2"] = "看法：每一列是一個單位。左半部：年初常來的人，年底變成怎樣（仍常來＋下滑＋淡出＝年初常來）。右半部：年初不常來的人有多少變常來、第二季以後才出現的人、一直不常來的人。"
    headers = ["單位", "年初常來", "年底仍常來", "其中近一個月都沒來", "年底變不常來（下滑）", "年底都沒來（淡出）",
               "仍常來比例", "年初不常來、年底變常來（回升）", "本區間第二季以後才開始出現", "一直不常來"]
    rows = []
    for u, ps in U.items():
        mv = movement(ps)
        rows.append([u, mv["base"], mv["kept"], mv["kept_away"], mv["down"], mv["fade"],
                     (round(mv["ret"], 3) if mv["ret"] is not None else None), mv["up"],
                     mv["tc"]["本區間第二季以後才開始出現"], mv["tc"]["一直不常來"]])
    end = table(ws, 4, headers, rows, [16, 10, 11, 12, 12, 12, 10, 14, 11, 11])
    for rr in range(5, end):
        ws.cell(row=rr, column=7).number_format = "0%"
    r = end + 1
    title(ws, r, "細項：年初每一級的人，年底各在哪一級（人數）")
    r += 1
    for u, ps in U.items():
        mv = movement(ps)
        title(ws, r, "【%s】" % u, 12)
        mrows = [["年初%s" % t0] + [mv["m"][t0][t1] for t1 in TIERS] + [sum(mv["m"][t0].values())] for t0 in TIERS]
        r = table(ws, r + 1, ["", "年底穩定", "年底普通", "年底不常來", "年底沒來", "合計"], mrows) + 1

    # 4 個人近一年軌跡（內部）
    ws = wb.create_sheet("個人近一年軌跡（內部）")
    ws["A1"] = "內部參考，不對外轉發。等級軌跡：穩＝穩定、普＝普通、少＝不常來、無＝沒來。"
    ws["A2"] = "類型：" + "；".join("%s＝%s" % kv for kv in TYPE_NOTE.items())
    aux_h = ["小排或家聚會各段週數"] if res["aux"] else []
    headers = ["所屬區", "兒童", "姓名", "群組"] + seg_h + ["等級軌跡", "近4週", "類型", "近一年出席", "最近出席日期"] + aux_h + ["近一年出席紀錄（左舊右新）"]
    prow = []
    for p in sorted(res["people"], key=lambda p: (p["kid"], order[p["area"]], tyo[p["type"]], -p["total"], p["name"])):
        gh = []
        if res["aux"]:
            gh = ["無法比對" if p.get("gh_segs") is None else "／".join(str(x) for x in p["gh_segs"])]
        prow.append([p["area"], "是" if p["kid"] else "", p["name"], p["group"]] + p["segs"] +
                    ["".join(TIER_CH[t] for t in p["tiers"]), p["r4"], p["type"], p["total"], p["last"].isoformat()] + gh +
                    ["".join("●" if x else "○" for x in p["seq"])])
    table(ws, 4, headers, prow, [10, 6, 20, 8] + [7] * nseg + [10, 7, 18, 9, 13] + [16] * len(aux_h) + [60])
    ws.freeze_panes = "D5"

    # 5 本區間第二季以後才開始出現（內部）
    ws = wb.create_sheet("本區間第二季以後才開始出現（內部）")
    ws["A1"] = "內部參考。本區間第一季（第1段）及更早週次都沒有出席、第二季以後才第一次出席者；可能是新人、遷入者或久未出席後回來的聖徒。"
    ws["A2"] = NEW_NOTE
    aux_h = ["出現後12週小排或家聚會"] if res["aux"] else []
    headers = ["所屬區", "兒童", "姓名", "群組", "第一次出席", "出現後12週出席", "近4週", "狀態"] + aux_h
    so = {st: i for i, st in enumerate(["曾常來但近一個月沒來", "有再來但未穩定", "已穩定連結", "只來1～2次", "觀察中"])}
    nrows = []
    for p in sorted([p for p in res["people"] if p["newcomer"]], key=lambda p: (p["kid"], order[p["area"]], so[p["new_status"]], p["first"], p["name"])):
        gh = []
        if res["aux"]:
            gh = ["—" if "new_gh" not in p else ("有" if p["new_gh"] else "無")]
        nrows.append([p["area"], "是" if p["kid"] else "", p["name"], p["group"], res["sundays"][p["first"]].isoformat(),
                      p["new_n"], p["r4"], p["new_status"]] + gh)
    table(ws, 4, headers, nrows, [10, 6, 20, 8, 13, 14, 8, 22, 18])
    ws.freeze_panes = "A5"

    # 6 主日×小排家聚會
    if res["aux"]:
        ws = wb.create_sheet("主日×小排家聚會")
        ws["A1"] = "年底主日等級 × 年底有無小排或家聚會紀錄（人數）。無法比對者不計。下滑、淡出者的小排家聚會情形見「近一年概況」下半部。"
        r = 3
        for u, ps in U.items():
            ps2 = [p for p in ps if p.get("gh_segs") is not None]
            title(ws, r, "【%s】" % u, 12)
            xr = []
            for t in TIERS:
                y = sum(1 for p in ps2 if p["tiers"][-1] == t and p["gh_segs"][-1])
                n = sum(1 for p in ps2 if p["tiers"][-1] == t and not p["gh_segs"][-1])
                xr.append(["年底主日" + t, y, n, y + n])
            r = table(ws, r + 1, ["", "有小排或家聚會", "沒有", "合計"], xr, [16, 16, 10, 10]) + 1

    # 7 排除及週次說明
    ws = wb.create_sheet("排除及週次說明")
    erows = [[a, "臨時點名（來訪、訪客、臨時、福音朋友）", v] for a, v in sorted(res["placeholder"].items())]
    erows += [[a, "全年0出席（可能為長期停用未處理名單）", v] for a, v in sorted(res["zero_year"].items())]
    r = table(ws, 1, ["分區", "排除原因", "人數"], erows, [12, 40, 8]) + 1
    for t, items in (("人數範圍", [SCOPE_NOTE]), ("週次判定", res["notes"]), ("資料待確認", res["issues"]), ("資料限制", LIMITS)):
        title(ws, r, t)
        for x in items or ["無"]:
            r += 1
            ws.cell(row=r, column=1, value=x).font = Font(size=12)
        r += 2
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and ("年初" in c.value or "年底" in c.value):
                    c.value = by_month(res, c.value)
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
    ap.add_argument("--group")
    ap.add_argument("--home")
    args = ap.parse_args()
    os.makedirs(args.outdir, exist_ok=True)
    res = build(args.report, args.group, args.home)
    tag = res["sundays"][-1].strftime("%Y%m%d")
    txt = summary_text(res)
    with open(os.path.join(args.outdir, "近一年出席觀察_%s.txt" % tag), "w", encoding="utf-8") as f:
        f.write(by_month(res, txt + INTERNAL_SEP + internal_text(res)))
    hl = highlights_text(res)
    with open(os.path.join(args.outdir, "近一年重點摘要_%s.txt" % tag), "w", encoding="utf-8") as f:
        f.write(hl)
    write_excel(res, os.path.join(args.outdir, "近一年出席分析_%s.xlsx" % tag), txt, hl)
    n_ex = sum(res["placeholder"].values()) + sum(res["zero_year"].values())
    print("分析模式：%s" % ("主日＋小排＋家聚會" if res["aux"] else "只有主日"))
    print("近一年主日%d次（%s～%s），分%d段，分段前%d週只用於趨勢" % (len(res["sundays"]), res["sundays"][0], res["sundays"][-1], res["nseg"], res["seg_start"]))
    for n in res["notes"]:
        print("週次：" + n)
    print("報表總人數%d；臨時點名%d；全年0出席%d；分析對象%d（兒童%d）；無法分區%d" % (
        res["total_rows"], sum(res["placeholder"].values()), sum(res["zero_year"].values()), len(res["people"]),
        sum(1 for p in res["people"] if p["kid"]), res["total_rows"] - n_ex - len(res["people"])))
    print("資料待確認%d筆" % len(res["issues"]))
    print(by_month(res, txt))
    print(hl)


if __name__ == "__main__":
    main()
