# -*- coding: utf-8 -*-
"""
參考優先關心名單｜獨立回測程式

用法：
    python3 verify_lists.py 主日報表.xlsx 輸出資料夾 [--zero-week include|exclude] [--group 小排報表.xlsx --home 家聚會報表.xlsx]

做法：
不引用 sunday_attendance.py 的任何程式碼，從原始報表重新讀取、重新計算，
再逐一核對分析程式產出的文字名單與 Excel：
  1. 每位聖徒的姓名、所屬區、類別（順位）是否正確
  2. 每類標題的人數是否等於實際列出的人數
  3. 名單順序是否符合規則
  4. 沒有重複、沒有遺漏、沒有不該出現的人（排除對象、空白類別）
  5. Excel 出席矩陣的每一格是否與原始報表一致，次數統計是否正確
  6. 每個類別標題下的白話說明與數字條件是否正確
  7. 三份報表模式：＊標記（近一季小排、家聚會都沒有紀錄）、說明行、Excel 標＊欄
     單檔模式：不得出現任何＊
  8. 跨區名單：青職（群組「青職」）與大專（群組「大學」，2026-11-11 新增）的人員、類別、註記、＊與人數
  9. 接近穩定名單（內部參考，2026-11-12 新增）：文字名單最下面與 Excel 工作表
 10. 主日沒來但有小排家聚會（內部名單，2026-11-12 起獨立）：三份報表模式時，文字名單最下面一段與 Excel 工作表
全部通過才輸出「回測通過」，任何一項不符即列出明細並以非零代碼結束。
"""
import re
import sys
import glob
import argparse
import datetime as dt
from collections import Counter, defaultdict

import openpyxl

AREA_ORDER = ["1小區", "2小區", "3小區", "4小區", "5小區", "6小區", "7小區", "青年區", "兒童專項"]
SKIP_WORDS = ("來訪", "訪客", "臨時", "福音朋友")
RANK_NAMES = {
    1: "出席穩定者本週缺席",
    2: "出席穩定者連續缺席兩週以上",
    3: "出席普通者近期缺席",
    4: "出席普通者連續缺席三週以上",
    5: "總出席低，但最近一個月來了一半以上",
    6: "總出席低，但最近一個月有來1次",
    7: "總出席低，原本隔週來，但最近一個月都沒來",
    8: "總出席低，原本很少來，且最近一個月都沒來",
    9: "過去不時出現，但近12週都沒來",
}
# 類別標題下兩行：白話說明、數字條件（數字條件外加全形括號）
RANK_DESC = {
    1: ("往常固定來，這週沒來", "（12週來9次以上，上週有來、這週沒來）"),
    2: ("往常固定來，但連續幾週沒來", "（12週來9次以上，已連續兩週以上沒來）"),
    3: ("往常還算常來，最近一兩週沒來", "（12週來6～8次，最近一兩週沒來）"),
    4: ("往常還算常來，但已有一段時間沒來", "（12週來6～8次，已連續三週以上沒來）"),
    5: ("原本不算常來或是新人，最近一個月來了一半以上", "（前8週來0～3次，近4週來2次以上；12週合計5次以下）"),
    6: ("原本不常來或是新人，最近一個月來了1次", "（前8週來0～4次，近4週來1次）"),
    7: ("原本大約隔週會來，但最近一個月都沒來", "（前8週來3～5次，近4週0次）"),
    8: ("原本很少來，最近一個月也都沒來", "（前8週來1～2次，近4週0次）"),
    9: ("以前偶爾會出現，但近三個月都沒見到", "（近12週0次，近一年內來過6次以上）"),
}
NAME_TO_RANK = {v: k for k, v in RANK_NAMES.items()}
MARK_RANKS = (5, 6)  # 這兩類若最近一次主日有來，姓名下一行需註記
MARK_TEXT = "（本週有出席）"
STAR = "＊"
STAR_NOTE = "＊表示近一季（同統計期間12週）小排、家聚會也都沒有紀錄"
# 跨區名單：（文字名單標題、群組欄的值）
CROSS = [("青職（跨區）", "青職"), ("大專（跨區）", "大學")]
# 內部參考：接近穩定名單、主日沒來但有小排家聚會（文字名單最下面，各以一行「＝」分隔；Excel 工作表）
INTERNAL_SEP = "\n" + "＝" * 20 + "\n\n"
NEAR_TITLE = "接近穩定名單｜內部參考，不對外轉發"
NEAR_SHEET = "接近穩定（內部）"
CONTACT_TITLE = "主日沒來但有小排家聚會｜內部參考，不對外轉發"
CONTACT_SHEET = "主日沒來但有小排家聚會（內部）"


def internal_parts(txt_path, aux, errors):
    parts = open(txt_path, encoding="utf-8").read().split(INTERNAL_SEP)
    want = 3 if aux else 2
    if len(parts) != want:
        errors.append("文字名單最下面的內部名單段數不符：應有%d段（接近穩定%s），實際%d段"
                      % (want - 1, "、主日沒來但有小排家聚會" if aux else "", len(parts) - 1))
    return parts


def check_area_block(label, text, title, exp_lines, total, errors):
    """核對內部名單：標題、總人數、各區標題與每一行。exp_lines：{區: [行]}（已依規則排序）。"""
    lines = text.splitlines()
    if not lines or lines[0] != title:
        errors.append("%s標題不符" % label)
    if "共%d人" % total not in lines:
        errors.append("%s總人數不符：應為%d人" % (label, total))
    got, cur, heads = defaultdict(list), None, {}
    for l in lines:
        m = re.match(r"^【(.+?)】（(\d+)人）$", l)
        if m:
            cur = m.group(1)
            heads[cur] = int(m.group(2))
        elif l.startswith("　") and cur:
            got[cur].append(l)
    order = [a for a in AREA_ORDER if exp_lines.get(a)]
    if list(heads) != order:
        errors.append("%s的區別或順序不符：應為 %s" % (label, "、".join(order) or "無"))
    for a in set(exp_lines) | set(got):
        if exp_lines.get(a, []) != got.get(a, []):
            errors.append("%s【%s】不符：應為 %s" % (label, a, "、".join(x.strip() for x in exp_lines.get(a, [])) or "無"))
        if a in heads and heads[a] != len(got.get(a, [])):
            errors.append("%s【%s】標題人數與列出人數不符" % (label, a))


def check_near(parts, wb, people, errors):
    """從重算結果核對接近穩定名單：12週6～8次、連續缺席未達3週（含兒童專項）。"""
    exp = [p for p in people.values() if 6 <= p["n"] <= 8 and p["miss"] <= 2]
    exp_lines = defaultdict(list)
    for p in sorted(exp, key=lambda p: (-p["n"], p["miss"], p["name"])):
        exp_lines[p["area"]].append("　%s（%s｜12週%d次｜%s）" % (p["name"], p["group"] or "無群組", p["n"],
                                                             "本週有來" if p["miss"] == 0 else "最近缺席%d週" % p["miss"]))
    if len(parts) < 2:
        errors.append("文字名單最下面缺少接近穩定名單（內部參考）")
        return 0
    check_area_block("接近穩定名單", parts[1], NEAR_TITLE, exp_lines, len(exp), errors)
    if NEAR_SHEET not in wb.sheetnames:
        errors.append("Excel 缺少「%s」工作表" % NEAR_SHEET)
    else:
        rows = [r for r in wb[NEAR_SHEET].iter_rows(min_row=5, values_only=True) if r[1]]
        g = sorted((r[0], str(r[1]).strip(), r[3], r[5]) for r in rows)
        e = sorted((p["area"], p["name"], p["n"], p["miss"]) for p in exp)
        if g != e:
            errors.append("Excel「%s」工作表與重算結果不符" % NEAR_SHEET)
    return len(exp)


def check_contact(parts, wb, exp_contact, errors):
    """主日沒來但有小排家聚會：近12週主日0次且近一年未達6次，但近一季小排或家聚會有紀錄者。"""
    order = {a: i for i, a in enumerate(AREA_ORDER)}
    ps = sorted(exp_contact, key=lambda e: (order.get(e["area"], 99), -(e["g4"] + e["h4"]), -(e["g12"] + e["h12"]), e["name"]))
    exp_lines = defaultdict(list)
    for e in ps:
        exp_lines[e["area"]].append("　%s（%s｜小排%d週、家聚會%d週｜近一個月小排%d週、家聚會%d週）" % (
            e["name"], e["group"] or "無群組", e["g12"], e["h12"], e["g4"], e["h4"]))
    if len(parts) < 3:
        errors.append("文字名單最下面缺少「主日沒來但有小排家聚會」名單")
    else:
        check_area_block("主日沒來但有小排家聚會", parts[2], CONTACT_TITLE, exp_lines, len(ps), errors)
    if CONTACT_SHEET not in wb.sheetnames:
        errors.append("Excel 缺少「%s」工作表" % CONTACT_SHEET)
    else:
        erows = list(wb[CONTACT_SHEET].iter_rows(min_row=4, values_only=True))
        got_e = [(r[0], str(r[1]).strip(), r[3], r[4], r[5], r[6]) for r in erows if r[1]]
        want_e = [(e["area"], e["name"], e["g12"], e["h12"], e["g4"], e["h4"]) for e in ps]
        if got_e != want_e:
            for k in sorted(set(want_e) - set(got_e)):
                errors.append("「%s」工作表漏列或次數錯：%s（%s）" % (CONTACT_SHEET, k[1], k[0]))
            for k in sorted(set(got_e) - set(want_e)):
                errors.append("「%s」工作表多列或次數錯：%s（%s）" % (CONTACT_SHEET, k[1], k[0]))
            if set(got_e) == set(want_e):
                errors.append("「%s」工作表排序不符" % CONTACT_SHEET)
    return len(ps)


# ---------- 一、從原始報表重新計算 ----------

def load_raw(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["出席"] if "出席" in wb.sheetnames else wb.worksheets[0]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    ref = None
    for r in grid[:5]:
        for v in r:
            m = re.search(r"製表日期\D*(\d{4})-(\d{1,2})-(\d{1,2})", str(v or ""))
            if m:
                ref = dt.date(*map(int, m.groups()))
    hi = next(i for i, r in enumerate(grid) if "姓名" in [str(x).strip() for x in r if x])
    head = [str(x).strip() if x else "" for x in grid[hi]]
    pos = {h: i for i, h in enumerate(head) if h}
    dates = []
    for i, h in enumerate(head):
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", h):
            dates.append((i, dt.date.fromisoformat(h) + dt.timedelta(days=6)))
    rows = [r for r in grid[hi + 1:] if r[pos["姓名"]]]
    return ref, pos, dates, rows


def recompute(path, zero_week):
    ref, pos, dates, rows = load_raw(path)
    totals = {i: sum(int(r[i] or 0) for r in rows) for i, _ in dates}
    # 重複週次欄：只保留有出席資料的那一欄（全空則保留第一欄）；兩欄以上都有資料即回測失敗
    grouped = defaultdict(list)
    for i, d in dates:
        grouped[d].append(i)
    keep = []
    for d, cols in grouped.items():
        filled = [i for i in cols if totals[i] > 0]
        if len(filled) > 1:
            print("【回測失敗】%s 主日有多個重複週次欄都有出席資料，需人工確認。" % d)
            sys.exit(1)
        keep.append(((filled or cols)[0], d))
    dates = sorted(keep, key=lambda x: x[1])
    past = [(i, d) for i, d in dates if d < ref]
    last = past[-1][1]
    window = [(i, d) for i, d in past if d > last - dt.timedelta(days=364)]
    nonzero = [(i, d) for i, d in window if totals[i] > 0]
    cut = nonzero[-12][1]
    usable = []
    for i, d in window:
        if totals[i] == 0 and d < cut:
            continue
        if totals[i] == 0 and zero_week != "include":
            continue
        usable.append((i, d))
    recent, before = usable[-12:], usable[:-12]

    people = {}
    dropped = []  # 近12週0次且一年內未達6次而排除者（姓名鍵、第幾筆、所屬區、姓名、群組）
    seen_rk = Counter()
    for r in rows:
        name = str(r[pos["姓名"]]).strip()
        big = str(r[pos["大區"]] or "").strip()
        small = str(r[pos["小區"]] or "").strip()
        rk = (name, big, small)
        seen_rk[rk] += 1
        occ = seen_rk[rk]
        grp = str(r[pos["群組"]] or "").strip() if "群組" in pos else ""
        blob = " ".join(str(r[pos[k]] or "") for k in ("姓名", "大區", "小區", "排", "群組") if k in pos)
        if any(w in blob for w in SKIP_WORDS):
            continue
        area = "青年區" if big == "會所青年區" else ("兒童專項" if big == "兒童專項" else small)
        rec = [1 if int(r[i] or 0) > 0 else 0 for i, _ in recent]
        old = sum(1 for i, _ in before if int(r[i] or 0) > 0)
        n = sum(rec)
        if n == 0 and old < 6:
            dropped.append((rk, occ, area, name, grp))
            continue
        if area not in AREA_ORDER:
            continue
        miss = len(rec) - (max(j for j, x in enumerate(rec) if x) + 1) if n else 12
        r4 = sum(rec[-4:])
        if n == 0:
            rank = 9
            miss = 12 + next((j for j, (i, _) in enumerate(reversed(before)) if int(r[i] or 0) > 0), len(before))
        elif n >= 9:
            rank = None if miss == 0 else (1 if miss == 1 else 2)
        elif n >= 6:
            rank = None if miss == 0 else (3 if miss <= 2 else 4)
        else:
            p8 = n - r4  # 前8週出席次數
            if r4 >= 2:
                rank = 5
            elif r4 == 1:
                rank = 6
            else:
                rank = 7 if p8 >= 3 else 8
        seen = [d for (i, d), x in zip(recent, rec) if x] or [d for i, d in before if int(r[i] or 0) > 0]
        key = (area, name)
        if key in people:
            people[key]["dup"] = True
        people[key] = dict(name=name, area=area, group=grp, rec=rec, n=n, miss=miss, r4=r4,
                           old=old, rank=rank, last=max(seen) if seen else None, dup=False, rk=rk, occ=occ)
    return ref, recent, people, dropped, seen_rk


def aux_weeks(path, recent):
    """小排／家聚會報表：回傳 {姓名鍵: [每筆在統計期間12週中有出席的週數]}。"""
    _, pos, dates, rows = load_raw(path)
    days = set(d for _, d in recent)
    out = defaultdict(list)
    for r in rows:
        rk = (str(r[pos["姓名"]]).strip(), str(r[pos["大區"]] or "").strip(), str(r[pos["小區"]] or "").strip())
        hit = set(d for i, d in dates if d in days and int(r[i] or 0) > 0)
        hit4 = set(d for d in hit if d in set(d for _, d in recent[-4:]))
        out[rk].append((len(hit), len(hit4)))
    return out


def judge(rk, occ, scount, gw, hw):
    """回傳 (小排近一季, 家聚會近一季, 小排近一個月, 家聚會近一個月) 或 'notfound' / 'mismatch'。"""
    if not gw.get(rk) and not hw.get(rk):
        return "notfound"
    if len(gw.get(rk, [])) != scount or len(hw.get(rk, [])) != scount:
        return "mismatch"
    g, h = gw[rk][occ - 1], hw[rk][occ - 1]
    return (g[0], h[0], g[1], h[1])


def exp_all(exp):
    return [p for lst in exp.values() for p in lst]


def expected_order(lst, rank):
    if rank == 1:
        k = lambda p: (-p["n"], p["name"])
    elif rank in (2, 3, 4):
        k = lambda p: (p["miss"], -p["n"], p["name"])
    elif rank in (5, 6):
        k = lambda p: (-p["r4"], -p["last"].toordinal(), p["name"])
    elif rank == 9:
        k = lambda p: (-p["old"], -p["last"].toordinal(), p["name"])
    else:
        k = lambda p: (-p["last"].toordinal(), p["name"])
    return [p["name"] for p in sorted(lst, key=k)]


# ---------- 二、讀取分析程式的產出 ----------

def parse_text(path):
    text = open(path, encoding="utf-8").read().split(INTERNAL_SEP)[0]  # 內部參考另行核對
    sections = {}
    cur, rank = None, None
    for line in text.splitlines():
        m = re.match(r"^(.+?)｜參考優先關心名單$", line)
        if m:
            cur, rank = m.group(1), None
            sections[cur] = {"ranks": [], "lists": {}, "counts": {}, "marks": defaultdict(list), "desc": {},
                             "stars": [], "starnote": 0}
            continue
        if line == STAR_NOTE and cur and rank is None:
            sections[cur]["starnote"] += 1
            continue
        m = re.match(r"^【(.+?)】(.*)（(\d+)人）$", line)
        if m and cur:
            rank = NAME_TO_RANK.get(m.group(1))
            if rank is None:
                sections[cur].setdefault("unknown", []).append(m.group(1))
                continue
            if m.group(2):
                sections[cur].setdefault("badhead", []).append(line)
            sections[cur]["ranks"].append(rank)
            sections[cur]["counts"][rank] = int(m.group(3))
            sections[cur]["lists"][rank] = []
            sections[cur]["desc"][rank] = []
            continue
        if line.startswith("　　") and cur and rank:
            if line[2:] != MARK_TEXT or not sections[cur]["lists"][rank]:
                sections[cur].setdefault("badmark", []).append(line)
            else:
                sections[cur]["marks"][rank].append(sections[cur]["lists"][rank][-1])
            continue
        if line.startswith("　") and cur and rank:
            nm = line[1:]
            if nm.startswith(STAR):
                nm = nm[1:]
                sections[cur]["stars"].append(nm)
            sections[cur]["lists"][rank].append(nm)
            continue
        if line.strip() and not set(line.strip()) <= {"—"} and cur and rank:
            if sections[cur]["lists"][rank]:
                sections[cur].setdefault("baddesc", []).append(line)
            else:
                sections[cur]["desc"][rank].append(line)
    return sections


def check_desc(label, got, errors):
    for b in got.get("badhead", []):
        errors.append("%s：類別標題格式不符「%s」" % (label, b))
    for b in got.get("baddesc", []):
        errors.append("%s：名單中出現不明文字「%s」" % (label, b))
    for rank, lines in got["desc"].items():
        exp = list(RANK_DESC[rank])
        if lines != exp:
            errors.append("%s【%s】說明行不符：應為「%s」；實際「%s」"
                          % (label, RANK_NAMES[rank], "／".join(exp), "／".join(lines) or "無"))


# ---------- 三、核對 ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("outdir")
    ap.add_argument("--zero-week", choices=["include", "exclude"])
    ap.add_argument("--group")
    ap.add_argument("--home")
    a = ap.parse_args()

    errors = []
    ref, recent, people, dropped, scount = recompute(a.report, a.zero_week)
    aux = bool(a.group and a.home)
    exp_contact = []
    if aux:
        gw, hw = aux_weeks(a.group, recent), aux_weeks(a.home, recent)
        for p in people.values():
            j = judge(p["rk"], p["occ"], scount[p["rk"]], gw, hw)
            p["aux"] = j
            p["star"] = bool(p["rank"]) and (j == "notfound" or (isinstance(j, tuple) and j[0] == 0 and j[1] == 0))
        for rk, occ, area, name, grp in dropped:
            j = judge(rk, occ, scount[rk], gw, hw)
            if isinstance(j, tuple) and (j[0] or j[1]):
                exp_contact.append(dict(area=area, name=name, group=grp, g12=j[0], h12=j[1], g4=j[2], h4=j[3]))
    else:
        for p in people.values():
            p["star"] = False

    txts = glob.glob(a.outdir + "/各區看望名單_*.txt")
    xlsxs = glob.glob(a.outdir + "/主日出席分析_*.xlsx")
    if len(txts) != 1 or len(xlsxs) != 1:
        print("【回測失敗】輸出資料夾中找不到唯一的文字名單或 Excel。")
        sys.exit(1)
    sec = parse_text(txts[0])

    for key, p in people.items():
        if p["dup"]:
            errors.append("原始報表同區重複姓名：%s（%s），請人工確認" % key)

    # 文字名單：各區
    checked = 0
    for area in AREA_ORDER:
        exp = defaultdict(list)
        for p in people.values():
            if p["area"] == area and p["rank"]:
                exp[p["rank"]].append(p)
        got = sec.get(area)
        if got is None:
            errors.append("%s：文字名單缺少整區" % area)
            continue
        for u in got.get("unknown", []):
            errors.append("%s：出現不明類別【%s】" % (area, u))
        for b in got.get("badmark", []):
            errors.append("%s：註記行格式不符「%s」" % (area, b))
        check_desc(area, got, errors)
        e_star = sorted(p["name"] for p in exp_all(exp) if p["star"])
        if sorted(got["stars"]) != e_star:
            errors.append("%s：＊標記不符：應為 %s；實際 %s" % (area, "、".join(e_star) or "無", "、".join(sorted(got["stars"])) or "無"))
        if got["starnote"] != (1 if aux else 0):
            errors.append("%s：＊說明行%s" % (area, "缺少" if aux else "不應出現"))
        for rank in range(1, 10):
            e_mark = sorted(p["name"] for p in exp.get(rank, []) if rank in MARK_RANKS and p["rec"][-1])
            g_mark = sorted(got["marks"].get(rank, []))
            if e_mark != g_mark:
                errors.append("%s【%s】本週有出席註記不符：應為 %s；實際 %s" % (area, RANK_NAMES[rank], "、".join(e_mark) or "無", "、".join(g_mark) or "無"))
        if got["ranks"] != sorted(got["ranks"]):
            errors.append("%s：類別順序錯誤 %s" % (area, got["ranks"]))
        for rank in range(1, 10):
            e_names = expected_order(exp.get(rank, []), rank)
            g_names = got["lists"].get(rank)
            if not e_names:
                if g_names is not None:
                    errors.append("%s：無人的類別【%s】不應列出" % (area, RANK_NAMES[rank]))
                continue
            if g_names is None:
                errors.append("%s：缺少類別【%s】（應有%d人）" % (area, RANK_NAMES[rank], len(e_names)))
                continue
            if got["counts"][rank] != len(g_names):
                errors.append("%s【%s】標題寫%d人，實際列出%d人" % (area, RANK_NAMES[rank], got["counts"][rank], len(g_names)))
            miss = set(e_names) - set(g_names)
            extra = set(g_names) - set(e_names)
            for nm in sorted(miss):
                p = people[(area, nm)]
                errors.append("%s【%s】漏列：%s（12週%d次、連缺%d週、近四週%d次）" % (area, RANK_NAMES[rank], nm, p["n"], p["miss"], p["r4"]))
            for nm in sorted(extra):
                p = people.get((area, nm))
                why = "應在【%s】" % RANK_NAMES[p["rank"]] if p and p["rank"] else "不應出現在名單"
                errors.append("%s【%s】多列：%s，%s" % (area, RANK_NAMES[rank], nm, why))
            if not miss and not extra and e_names != g_names:
                errors.append("%s【%s】排序不符：應為 %s" % (area, RANK_NAMES[rank], "、".join(e_names)))
            if len(set(g_names)) != len(g_names):
                errors.append("%s【%s】有重複姓名" % (area, RANK_NAMES[rank]))
            checked += len(g_names)

    # 文字名單：跨區（青職、大專）
    cross_checked = {}
    for title, grp_val in CROSS:
        lb = title.replace("（跨區）", "")
        yz = sec.get(title)
        if yz is None:
            errors.append("%s：文字名單缺少" % title)
            continue
        check_desc(title, yz, errors)
        e_star = sorted("%s（%s）" % (p["name"], p["area"]) for p in people.values() if p["group"] == grp_val and p["rank"] and p["star"])
        if sorted(yz["stars"]) != e_star:
            errors.append("%s：＊標記不符：應為 %s；實際 %s" % (lb, "、".join(e_star) or "無", "、".join(sorted(yz["stars"])) or "無"))
        if yz["starnote"] != (1 if aux else 0):
            errors.append("%s：＊說明行%s" % (lb, "缺少" if aux else "不應出現"))
        if yz["ranks"] != sorted(yz["ranks"]):
            errors.append("%s：類別順序錯誤 %s" % (lb, yz["ranks"]))
        n_chk = 0
        for rank in range(1, 10):
            e = sorted("%s（%s）" % (p["name"], p["area"]) for p in people.values() if p["group"] == grp_val and p["rank"] == rank)
            g = sorted(yz["lists"].get(rank, []))
            if e != g:
                errors.append("%s【%s】不符：應有 %s；實際 %s" % (lb, RANK_NAMES[rank], "、".join(e) or "無", "、".join(g) or "無"))
            if not e and rank in yz["lists"]:
                errors.append("%s：無人的類別【%s】不應列出" % (lb, RANK_NAMES[rank]))
            em = sorted("%s（%s）" % (p["name"], p["area"]) for p in people.values()
                        if p["group"] == grp_val and p["rank"] == rank and rank in MARK_RANKS and p["rec"][-1])
            gm = sorted(yz["marks"].get(rank, []))
            if em != gm:
                errors.append("%s【%s】本週有出席註記不符：應為 %s；實際 %s" % (lb, RANK_NAMES[rank], "、".join(em) or "無", "、".join(gm) or "無"))
            if rank in yz["counts"] and yz["counts"][rank] != len(yz["lists"][rank]):
                errors.append("%s【%s】標題人數與列出人數不符" % (lb, RANK_NAMES[rank]))
            n_chk += len(g)
        cross_checked[lb] = n_chk

    # Excel：出席矩陣逐格核對
    wb = openpyxl.load_workbook(xlsxs[0], data_only=True)
    ws = wb["12次出席矩陣"]
    rows = list(ws.iter_rows(values_only=True))
    head = list(rows[0])
    exp_dates = [d.strftime("%m/%d") for _, d in recent]
    if head[3:15] != exp_dates:
        errors.append("Excel 出席矩陣的12個主日日期不符：應為 %s" % "、".join(exp_dates))
    col = {h: i for i, h in enumerate(head)}
    seen_keys = set()
    for r in rows[1:]:
        key = (r[0], str(r[1]).strip())
        seen_keys.add(key)
        p = people.get(key)
        if p is None:
            errors.append("Excel 矩陣多出不應納入的人：%s（%s）" % (key[1], key[0]))
            continue
        cells = [int(x or 0) for x in r[3:15]]
        if cells != p["rec"]:
            errors.append("Excel 矩陣出席紀錄與原始報表不符：%s（%s）" % (key[1], key[0]))
        if r[col["12週出席"]] != p["n"]:
            errors.append("Excel 12週出席次數錯誤：%s 寫%s，應為%d" % (key[1], r[col["12週出席"]], p["n"]))
        if r[col["近四週出席"]] != p["r4"]:
            errors.append("Excel 近四週出席次數錯誤：%s" % key[1])
        if r[col["連續缺席週數"]] != p["miss"]:
            errors.append("Excel 連續缺席週數錯誤：%s 寫%s，應為%d" % (key[1], r[col["連續缺席週數"]], p["miss"]))
        if "近12週以前一年內出席" in col and r[col["近12週以前一年內出席"]] != p["old"]:
            errors.append("Excel 近一年出席次數錯誤：%s" % key[1])
        if aux:
            if "標＊" not in col:
                errors.append("Excel 矩陣缺少「標＊」欄")
            elif (r[col["標＊"]] == STAR) != p["star"]:
                errors.append("Excel 矩陣標＊錯誤：%s（%s）" % (key[1], key[0]))
            elif isinstance(p["aux"], tuple) and (r[col["小排近一季"]], r[col["家聚會近一季"]]) != p["aux"][:2]:
                errors.append("Excel 矩陣小排／家聚會次數錯誤：%s（%s）" % (key[1], key[0]))
        elif "標＊" in col:
            errors.append("Excel 單檔模式不應有「標＊」欄")
        label = str(r[col["類別"]])
        if p["rank"]:
            if not label.startswith("第%d順位" % p["rank"]):
                errors.append("Excel 類別錯誤：%s 寫「%s」，應為第%d順位" % (key[1], label, p["rank"]))
        elif not label.startswith("不列入"):
            errors.append("Excel 類別錯誤：%s 應為不列入" % key[1])
    for key in set(people) - seen_keys:
        errors.append("Excel 矩陣漏列：%s（%s）" % (key[1], key[0]))

    # Excel：全區看望總表
    ws = wb["全區看望總表"]
    trows = list(ws.iter_rows(values_only=True))[1:]
    got_t = Counter((r[2], str(r[3]).strip(), r[0]) for r in trows)
    exp_t = Counter((p["area"], p["name"], p["rank"]) for p in people.values() if p["rank"])
    if got_t != exp_t:
        for k in (exp_t - got_t):
            errors.append("全區看望總表漏列：%s（%s，第%d順位）" % (k[1], k[0], k[2]))
        for k in (got_t - exp_t):
            errors.append("全區看望總表多列或順位錯：%s（%s，第%s順位）" % (k[1], k[0], k[2]))
    ranks_seq = [r[0] for r in trows]
    if ranks_seq != sorted(ranks_seq):
        errors.append("全區看望總表未依順位排列")

    # 內部名單：接近穩定、主日沒來但有小排家聚會
    parts = internal_parts(txts[0], aux, errors)
    n_near = check_near(parts, wb, people, errors)
    n_contact = 0
    if aux:
        n_contact = check_contact(parts, wb, exp_contact, errors)
    elif CONTACT_SHEET in wb.sheetnames:
        errors.append("單檔模式不應有「%s」工作表" % CONTACT_SHEET)

    # 重點摘要（2026-11-13 新增）：名單與人數
    hls = glob.glob(a.outdir + "/關心名單重點摘要_*.txt")
    if len(hls) != 1:
        errors.append("找不到唯一的關心名單重點摘要檔")
    else:
        hl = open(hls[0], encoding="utf-8").read()
        # 登載示警的區不列（最近一次主日全區0、前11次平均大於0）
        warn = set()
        for ar in AREA_ORDER:
            ps_ = [p for p in people.values() if p["area"] == ar]
            if ps_ and sum(p["rec"][-1] for p in ps_) == 0 and sum(sum(p["rec"][:-1]) for p in ps_) > 0:
                warn.add(ar)

        def ordered(cond):
            out = []
            for rank in range(1, 10):
                for ar in AREA_ORDER:
                    if ar in warn:
                        continue
                    lst = [p for p in people.values() if p["area"] == ar and p["rank"] == rank and cond(p)]
                    out += [(STAR if p["star"] else "") + n for n in expected_order(lst, rank) for p in [people[(ar, n)]]]
            return out

        def names_after(marker):
            if marker not in hl:
                return None
            blk = hl.split(marker, 1)[1].split("\n〔", 1)[0].split("\n【", 1)[0]
            return [re.match(r"^　(.+)（[^（｜]+｜", l).group(1) for l in blk.splitlines() if l.startswith("　") and l != "　無"]

        def section(title):
            return hl.split(title, 1)[1].split("\n【", 1)[0] if title in hl else ""
        checks = [("【二、", ordered(lambda p: p["rank"] in (1, 2))),
                  ("【三、", ordered(lambda p: p["rank"] == 4 or (aux and p["rank"] == 7 and p["star"])))]
        if aux:
            checks.append(("〔可能已斷了聯繫", ordered(lambda p: p["rank"] == 9 and p["star"])))
            order_c = {x: i for i, x in enumerate(AREA_ORDER)}
            inv = sorted([e for e in exp_contact if e["g4"] + e["h4"] >= 2],
                         key=lambda e: (order_c.get(e["area"], 99), -(e["g4"] + e["h4"]), -(e["g12"] + e["h12"]), e["name"]))
            checks.append(("【五、", [e["name"] for e in inv]))
        for marker, want in checks:
            if marker.startswith("〔"):
                got = names_after(marker)
            else:
                sec = section(marker)
                blk = sec.split("〔", 1)[1] if "〔" in sec else ""
                blk = blk.split("\n〔", 1)[0]
                got = [re.match(r"^　(.+)（[^（｜]+｜", l).group(1) for l in blk.splitlines() if l.startswith("　") and l != "　無"]
            if got != want:
                errors.append("重點摘要%s名單不符：應為 %s；實際 %s" % (marker, "、".join(want) or "無", "、".join(got or []) or "無"))
        listed = [p for p in people.values() if p["rank"] and p["area"] not in warn]
        want_line = "・列入關心名單共%d人：原本常來的人缺席（第1～4順位）%d人、原本不常來（第5～8順位）%d人、近三個月都沒見到（第9順位）%d人。" % (
            len(listed), sum(1 for p in listed if p["rank"] <= 4), sum(1 for p in listed if 5 <= p["rank"] <= 8), sum(1 for p in listed if p["rank"] == 9))
        if want_line not in hl:
            errors.append("重點摘要第一點人數不符：應為「%s」" % want_line)
        back = sum(1 for p in people.values() if p["area"] not in warn and p["rank"] in MARK_RANKS and p["rec"][-1])
        if "本週有出席」）%d人" % back not in hl:
            errors.append("重點摘要第四點「本週有出席」人數不符：應為%d人" % back)
        if "接近穩定（12週來6～8次，再多來一兩次就達到穩定）%d人" % n_near not in hl:
            errors.append("重點摘要第四點接近穩定人數不符")

    total_listed = sum(1 for p in people.values() if p["rank"])
    print("回測基準：製表日期 %s；12次主日 %s～%s" % (ref, recent[0][1], recent[-1][1]))
    print("重新計算：有效名單 %d 人，列入看望順位 %d 人；文字名單核對 %d 筆" % (len(people), total_listed, checked))
    print("接近穩定名單（內部參考）核對：%d 人" % n_near)
    print("跨區名單核對：" + "、".join("%s %d 筆" % (k, v) for k, v in cross_checked.items()))
    if aux:
        print("三份報表模式：標＊ %d 人；主日沒來但有小排家聚會 %d 人" % (sum(1 for p in people.values() if p["rank"] and p["star"]), n_contact))
    if errors:
        print("【回測失敗】共 %d 項不符：" % len(errors))
        for e in errors:
            print("・" + e)
        sys.exit(1)
    print("【回測通過】姓名、所屬區、類別、說明行、排序、人數、出席次數與矩陣逐格皆與原始報表一致。")


if __name__ == "__main__":
    main()
