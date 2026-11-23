# -*- coding: utf-8 -*-
"""
近一年出席分析｜規則 V1.1 回測程式（不引用分析程式，從原始報表重新計算後核對輸出）

用法：python3 verify_annual.py 主日報表.xlsx 輸出資料夾 [--group 小排報表.xlsx --home 家聚會報表.xlsx]
2026-11-13 修訂：「近12週主日都沒來，但小排或家聚會有來」不限近一年主日次數。
"""
import re
import os
import sys
import argparse
import datetime as dt
from collections import Counter, defaultdict

import openpyxl

AREAS = ["1小區", "2小區", "3小區", "4小區", "5小區", "6小區", "7小區", "青年區", "兒童專項"]
ADULT = AREAS[:-1]
KW = ["來訪", "訪客", "臨時", "福音朋友"]
KIDG = ("學齡前", "國小")
CROSS = {"青職": "青職（跨區）", "大學": "大專（跨區）"}
LV = ["穩定", "普通", "不常來", "沒來"]
CH = {"穩定": "穩", "普通": "普", "不常來": "少", "沒來": "無"}
TYPES = ["留住", "常來但近一個月沒來", "下滑", "淡出", "回升", "本區間第二季以後才開始出現", "一直不常來"]

errors = []


def err(msg):
    errors.append(msg)


def lv(n):
    return "穩定" if n >= 9 else "普通" if n >= 6 else "不常來" if n else "沒來"


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["出席"] if "出席" in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    ref = None
    for r in rows[:10]:
        for v in r:
            m = re.search(r"製表日期\s*[:：]\s*(\d{4})-(\d{1,2})-(\d{1,2})", str(v or ""))
            if m:
                ref = dt.date(*map(int, m.groups()))
    h = next(i for i, r in enumerate(rows[:10]) if "姓名" in [str(x).strip() if x else "" for x in r])
    head = [str(x).strip() if x else "" for x in rows[h]]
    idx = {k: head.index(k) for k in ("姓名", "大區", "小區", "群組") if k in head}
    dates = defaultdict(list)
    for j, x in enumerate(head):
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", x):
            dates[dt.date.fromisoformat(x)].append(j)
    return ref, idx, dates, [r for r in rows[h + 1:] if r and r[idx["姓名"]]]


def block(cells, title):
    return next((n for n, r in enumerate(cells) if r[0] == title), None)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("outdir")
    ap.add_argument("--group")
    ap.add_argument("--home")
    a = ap.parse_args()
    aux = bool(a.group and a.home)

    ref, idx, dates, body = load(a.report)
    weeks = []
    for mon in sorted(dates):
        sun = mon + dt.timedelta(days=6)
        if sun >= ref:
            continue
        cols = [j for j in dates[mon] if any(int(r[j] or 0) for r in body)]
        if cols:
            weeks.append((sun, cols[0]))
    last = weeks[-1][0]
    yw = [(s, j) for s, j in weeks if (last - s).days < 364]
    early = [j for s, j in weeks if (last - s).days >= 364]
    W = len(yw)
    K = min(4, W // 12)
    off = W - 12 * K

    def area(r):
        b = (r[idx["大區"]] or "").strip()
        s = (r[idx["小區"]] or "").strip()
        if b == "會所青年區":
            return "青年區"
        if b == "兒童專項":
            return "兒童專項"
        return s if s in AREAS else None

    P = []
    Z = []  # 近一年主日0次者
    place_w = [0] * W
    seen = Counter()
    for r in body:
        k3 = (str(r[idx["姓名"]]).strip(), str(r[idx["大區"]] or "").strip(), str(r[idx["小區"]] or "").strip())
        seen[k3] += 1
        if any(w in " ".join(str(r[c] or "") for c in idx.values()) for w in KW):
            for i, (s, j) in enumerate(yw):
                place_w[i] += int(r[j] or 0)
            continue
        v = [int(bool(int(r[j] or 0))) for s, j in yw]
        ar = area(r)
        if not any(v) and ar is not None:
            Z.append(dict(k=k3 + (seen[k3],), name=k3[0], area=ar, grp=(r[idx["群組"]] if "群組" in idx else "") or ""))
        if not any(v) or ar is None:
            continue
        grp = (r[idx["群組"]] if "群組" in idx else "") or ""
        sc = [sum(v[off + 12 * i: off + 12 * i + 12]) for i in range(K)]
        ls = [lv(x) for x in sc]
        f = v.index(1)
        r4 = sum(v[-4:])
        new = f >= off + 12 and not any(int(r[j] or 0) for j in early)
        often = ("穩定", "普通")
        if new:
            t = "本區間第二季以後才開始出現"
        elif ls[-1] in often and r4 == 0:
            t = "常來但近一個月沒來"
        elif ls[0] in often:
            t = "留住" if ls[-1] in often else "下滑" if ls[-1] == "不常來" else "淡出"
        else:
            t = "回升" if ls[-1] in often else "一直不常來"
        d = dict(k=k3 + (seen[k3],), name=k3[0], area=ar, grp=grp, kid=ar == "兒童專項" or grp in KIDG,
                 v=v, sc=sc, ls=ls, t=t, new=new, f=f, r4=r4)
        if new:
            d["nn"] = sum(v[f:f + 12])
            if f + 12 > W:
                d["ns"] = "觀察中"
            elif d["nn"] >= 6:
                d["ns"] = "已穩定連結" if r4 else "曾常來但近一個月沒來"
            else:
                d["ns"] = "有再來但未穩定" if d["nn"] >= 3 else "只來1～2次"
        P.append(d)

    if aux:
        recs, cnts = [], []
        for path in (a.group, a.home):
            r2, i2, d2, b2 = load(path)
            if r2 != ref:
                err("%s 製表日期不同" % path)
            c, rec = Counter(), {}
            for r in b2:
                k3 = (str(r[i2["姓名"]]).strip(), str(r[i2["大區"]] or "").strip(), str(r[i2["小區"]] or "").strip())
                c[k3] += 1
                rec[k3 + (c[k3],)] = [int(any(int(r[j] or 0) for j in d2.get(s - dt.timedelta(days=6), []))) for s, _ in yw]
            recs.append(rec)
            cnts.append(c)
        for p in P:
            k3 = p["k"][:3]
            if cnts[0][k3] == 0 and cnts[1][k3] == 0:
                p["g"] = [0] * W
            elif cnts[0][k3] != seen[k3] or cnts[1][k3] != seen[k3]:
                p["g"] = None
            else:
                p["g"] = [x | y for x, y in zip(recs[0][p["k"]], recs[1][p["k"]])]
            if p["g"] is not None:
                p["gs"] = [sum(p["g"][off + 12 * i: off + 12 * i + 12]) for i in range(K)]
        NS = []
        # 2026-11-13 修訂：不限近一年主日次數
        cand = [dict(z, sy=0) for z in Z] + [dict(k=p["k"], name=p["name"], area=p["area"], grp=p["grp"], sy=sum(p["v"]))
                                             for p in P if p["sc"][-1] == 0]
        for z in cand:
            k3 = z["k"][:3]
            if cnts[0][k3] != seen[k3] or cnts[1][k3] != seen[k3]:
                continue
            g, h = recs[0][z["k"]], recs[1][z["k"]]
            z.update(gy=sum(g), hy=sum(h), gq=sum(g[off + 12 * (K - 1):]), hq=sum(h[off + 12 * (K - 1):]))
            if z["gq"] or z["hq"]:
                NS.append(z)
        ao = {x: i for i, x in enumerate(AREAS)}
        NS.sort(key=lambda z: (ao[z["area"]], -(z["gq"] + z["hq"]), -(z["gy"] + z["hy"]), z["name"]))

    U = {"全會所": [p for p in P if not p["kid"]]}
    for ar in ADULT:
        U[ar] = [p for p in P if p["area"] == ar and not p["kid"]]
    U["兒童"] = [p for p in P if p["kid"]]
    for g, t in CROSS.items():
        U[t] = [p for p in P if p["grp"] == g]

    def mv(ps):
        base = [p for p in ps if p["ls"][0] in ("穩定", "普通")]
        kept = [p for p in base if p["ls"][-1] in ("穩定", "普通")]
        return dict(base=len(base), kept=len(kept), away=sum(1 for p in kept if p["r4"] == 0),
                    down=sum(1 for p in base if p["ls"][-1] == "不常來"), fade=sum(1 for p in base if p["ls"][-1] == "沒來"),
                    up=sum(1 for p in ps if not p["new"] and p["ls"][0] in ("不常來", "沒來") and p["ls"][-1] in ("穩定", "普通")))

    def avg(ps, i):
        return sum(p["sc"][i] for p in ps) / 12

    def contrib(ps):
        return {t: sum(p["sc"][-1] - p["sc"][0] for p in ps if p["t"] == t) / 12 for t in TYPES}

    tag = last.strftime("%Y%m%d")
    xl = os.path.join(a.outdir, "近一年出席分析_%s.xlsx" % tag)
    tx = os.path.join(a.outdir, "近一年出席觀察_%s.txt" % tag)
    if not (os.path.exists(xl) and os.path.exists(tx)):
        print("【回測失敗】找不到輸出檔")
        sys.exit(1)
    wb = openpyxl.load_workbook(xl)
    def mspan(i):
        return "%d～%d月" % (yw[off + 12 * i][0].month, yw[off + 12 * i + 11][0].month)
    A, B = mspan(0), mspan(K - 1)

    def bm(t):
        return t.replace("年初", A).replace("年底", B)
    full = open(tx, encoding="utf-8").read()
    if "年初" in full or "年底" in full:
        err("文字摘要仍有「年初」「年底」字樣，未換成實際月份")
    sep = "\n" + "＝" * 20 + "\n\n"
    if full.count(sep) != 1:
        err("文字摘要缺少內部名單分隔線")
    txt, internal = (full.split(sep) + [""])[:2]
    left = [ws.title for ws in wb.worksheets for row in ws.iter_rows(values_only=True) for x in row
            if isinstance(x, str) and ("年初" in x or "年底" in x)]
    if left:
        err("Excel 仍有「年初」「年底」字樣：%s" % "、".join(sorted(set(left))))
    gh_last = (lambda p: ["無法比對" if p.get("gs") is None else p["gs"][-1]]) if aux else (lambda p: [])

    # 1 每週趨勢
    rows = [r for r in wb["每週趨勢"].iter_rows(min_row=3, values_only=True) if r[0] and re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(r[0]))]
    if len(rows) != W:
        err("每週趨勢週數 %d ≠ %d" % (len(rows), W))
    tot = []
    for i, (s, j) in enumerate(yw):
        per = [sum(p["v"][i] for p in U[ar]) for ar in ADULT]
        tot.append(sum(p["v"][i] for p in U["全會所"]))
        if tot[-1] != sum(per):
            err("每週趨勢 %s 各區加總≠全會所" % s)
        exp = (s.isoformat(), "分段前" if i < off else "第%d段" % ((i - off) // 12 + 1), tot[i],
               round(sum(tot[i - 3:i + 1]) / 4, 1) if i >= 3 else None, *per, sum(p["v"][i] for p in U["兒童"]), place_w[i])
        if i < len(rows) and tuple(rows[i][:len(exp)]) != exp:
            err("每週趨勢 %s：輸出 %s，重算 %s" % (s, rows[i][:len(exp)], exp))

    # 2 個人近一年軌跡
    out = Counter(tuple(r) for r in wb["個人近一年軌跡（內部）"].iter_rows(min_row=5, values_only=True) if r[0])
    exp = Counter()
    for p in P:
        g = (["無法比對" if p.get("gs") is None else "／".join(map(str, p["gs"]))] if aux else [])
        exp[tuple([p["area"], "是" if p["kid"] else None, p["name"], p["grp"]] + p["sc"] +
                  ["".join(CH[x] for x in p["ls"]), p["r4"], p["t"], sum(p["v"]),
                   max(s for (s, _), x in zip(yw, p["v"]) if x).isoformat()] + g + ["".join("●" if x else "○" for x in p["v"])])] += 1
    if out != exp:
        err("個人近一年軌跡不符（%d／%d 筆），例：%s" % (sum(out.values()), sum(exp.values()), list((exp - out).elements())[:2] + list((out - exp).elements())[:2]))

    # 3 年初年底對照
    cells = list(wb["前後季對照"].iter_rows(values_only=True))
    summ = {r[0]: r for r in cells[4:4 + len(U)]}
    for u, ps in U.items():
        m = mv(ps)
        r = summ.get(u)
        tc = Counter(p["t"] for p in ps)
        want = [m["base"], m["kept"], m["away"], m["down"], m["fade"], round(m["kept"] / m["base"], 3) if m["base"] else None,
                m["up"], tc["本區間第二季以後才開始出現"], tc["一直不常來"]]
        if r is None or list(r[1:10]) != want:
            err("年初年底對照 %s：輸出 %s，重算 %s" % (u, r and list(r[1:10]), want))
        i = block(cells, "【%s】" % u)
        if i is None:
            err("年初年底對照細項缺少 %s" % u)
            continue
        for n, x in enumerate(LV):
            got = list(cells[i + 2 + n][1:5])
            want = [sum(1 for p in ps if p["ls"][0] == x and p["ls"][-1] == y) for y in LV]
            if got != want:
                err("年初年底對照細項 %s 年初%s：%s ≠ %s" % (u, x, got, want))

    # 4 近一年概況下半部（含姓名）
    cells = list(wb["近一年概況"].iter_rows(values_only=True))
    i1 = next((n for n, r in enumerate(cells) if str(r[0] or "").startswith("一、下滑")), None)
    i2 = next((n for n, r in enumerate(cells) if str(r[0] or "").startswith("二、各區增減")), None)
    i3 = next((n for n, r in enumerate(cells) if str(r[0] or "").startswith("三、人數減少")), None)
    if None in (i1, i2, i3):
        err("近一年概況缺少內部名單段落")
    else:
        got = Counter(tuple(r[:4 + K + 1 + (1 if aux else 0)]) for r in cells[i1 + 3:i2] if r[0] and r[1])
        want = Counter(tuple([("兒童" if p["kid"] and p["area"] != "兒童專項" else "") + p["area"], p["name"], p["grp"], p["t"]] + p["sc"] + [p["r4"]] + gh_last(p))
                       for p in P if p["t"] in ("下滑", "淡出", "常來但近一個月沒來"))
        if got != want:
            err("近一年概況 下滑淡出名單不符（%d／%d）" % (sum(got.values()), sum(want.values())))
        dec = [ar for ar in ADULT if avg(U[ar], 0) and (avg(U[ar], K - 1) - avg(U[ar], 0)) / avg(U[ar], 0) <= -0.10]
        units4 = ADULT + ["兒童", "青職（跨區）", "大專（跨區）"]
        rows2 = {r[0]: r for r in cells[i2 + 3:i3] if r[0] in units4}
        for ar in units4:
            c = contrib(U[ar])
            r = rows2.get(ar)
            want = [round(avg(U[ar], 0), 1), round(avg(U[ar], K - 1), 1), round(avg(U[ar], K - 1) - avg(U[ar], 0), 1)] + [round(c[t], 1) for t in TYPES]
            if r is None or list(r[1:4 + len(TYPES)]) != want:
                err("近一年概況 增減拆解 %s 不符：%s ≠ %s" % (ar, r and list(r[1:4 + len(TYPES)]), want))
            if abs(sum(c.values()) - (avg(U[ar], K - 1) - avg(U[ar], 0))) > 1e-9:
                err("增減拆解加總不等於總增減：%s" % ar)
        i4 = next((n for n, r in enumerate(cells) if str(r[0] or "").startswith("四、")), len(cells))
        got = Counter(tuple(r[:2]) for r in cells[i3 + 3:i4] if r[0] in ADULT and r[1])
        want = Counter((ar, p["name"]) for ar in dec for p in U[ar] if p["sc"][0] - p["sc"][-1] >= 3)
        if got != want:
            err("近一年概況 人數減少區名單不符（%d／%d）" % (sum(got.values()), sum(want.values())))

    # 5 本區間第二季以後才開始出現
    out = Counter(tuple(r[:8]) for r in wb["本區間第二季以後才開始出現（內部）"].iter_rows(min_row=5, values_only=True) if r[0])
    exp = Counter((p["area"], "是" if p["kid"] else None, p["name"], p["grp"], yw[p["f"]][0].isoformat(), p["nn"], p["r4"], p["ns"]) for p in P if p["new"])
    if out != exp:
        err("本區間第二季以後才開始出現名單不符（%d／%d）" % (sum(out.values()), sum(exp.values())))

    # 6 主日×小排家聚會
    if aux:
        cells = list(wb["主日×小排家聚會"].iter_rows(values_only=True))
        for u, ps in U.items():
            i = block(cells, "【%s】" % u)
            if i is None:
                err("主日×小排家聚會缺少 %s" % u)
                continue
            ps2 = [p for p in ps if p.get("gs") is not None]
            for n, x in enumerate(LV):
                y = sum(1 for p in ps2 if p["ls"][-1] == x and p["gs"][-1])
                z = sum(1 for p in ps2 if p["ls"][-1] == x and not p["gs"][-1])
                if list(cells[i + 2 + n][1:4]) != [y, z, y + z]:
                    err("主日×小排家聚會 %s %s 不符" % (u, x))

    # 7 文字摘要：重點數字、訊號、不含姓名
    m = mv(U["全會所"])
    for want in [bm("年初常來共%d人") % m["base"], "仍常來%d人（%.0f%%），其中近一個月都沒來%d人" % (m["kept"], m["kept"] / m["base"] * 100, m["away"]),
                 "變不常來（下滑）%d人" % m["down"], "都沒來（淡出）%d人" % m["fade"]]:
        if want not in txt:
            err("摘要缺少：" + want)
    ret_all = m["kept"] / m["base"]
    sig = []
    for ar in ADULT:
        ps = U[ar]
        a0, a1 = avg(ps, 0), avg(ps, K - 1)
        if a0 and (a1 - a0) / a0 <= -0.10:
            c = contrib(ps)
            t = min(c, key=lambda k: c[k])
            sig.append("%s：年底平均每週%.1f人，比年初的%.1f人減少%.0f%%；減少最多來自「%s」（%+.1f人）" % (ar, a1, a0, (a0 - a1) / a0 * 100, t, c[t]))
        mm = mv(ps)
        if mm["base"] >= 10 and mm["kept"] / mm["base"] <= ret_all - 0.10:
            sig.append("%s：年初常來的人年底仍常來的比例%.0f%%（%d／%d），比全會所%.0f%%低%.0f個百分點" % (
                ar, mm["kept"] / mm["base"] * 100, mm["kept"], mm["base"], ret_all * 100, (ret_all - mm["kept"] / mm["base"]) * 100))
        if mm["base"] >= 10 and mm["fade"] / mm["base"] >= 0.15:
            sig.append("%s：年初常來的%d人中，有%d人（%.0f%%）年底都沒來" % (ar, mm["base"], mm["fade"], mm["fade"] / mm["base"] * 100))
    if aux:
        fade = [p for p in U["全會所"] if p["t"] == "淡出" and p.get("gs") is not None]
        sig.append("全會所：年底主日都沒來的%d人中，有%d人年底小排、家聚會也都沒有紀錄" % (len(fade), sum(1 for p in fade if not p["gs"][-1])))
    sec = txt.split("【待了解訊號】")[1].split("【資料限制】")[0] if "【待了解訊號】" in txt else None
    got = re.findall(r"^・(.+)$", sec, re.M) if sec is not None else None
    sig = [bm(x) for x in sig]
    if got is None or got != sig:
        err("待了解訊號不符：輸出 %s；重算 %s" % (got, sig))
    names = [p["name"] for p in P if len(p["name"]) >= 2 and p["name"] in txt]
    if names:
        err("摘要出現姓名：%s" % "、".join(names[:5]))

    # 8 文字摘要內部名單
    sec1 = internal.split("【一、")[1].split("【二、")[0] if "【一、" in internal and "【二、" in internal else ""
    sec2 = internal.split("【二、")[1].split("【三、")[0] if "【二、" in internal else ""
    sec_up = internal.split("【三、")[1].split("【四、")[0] if "【三、" in internal else ""
    got = Counter(re.findall(r"^　(\S+?)（[^｜]*｜(\S+?)｜", sec1, re.M))
    want = Counter((p["name"], p["t"]) for p in P if p["t"] in ("下滑", "淡出", "常來但近一個月沒來"))
    if got != want:
        err("文字摘要內部名單一不符（%d／%d）" % (sum(got.values()), sum(want.values())))
    dec = [ar for ar in ADULT if avg(U[ar], 0) and (avg(U[ar], K - 1) - avg(U[ar], 0)) / avg(U[ar], 0) <= -0.10]
    got = Counter(re.findall(r"^　(\S+?)（", sec2, re.M))
    want = Counter(p["name"] for ar in dec for p in U[ar] if p["sc"][0] - p["sc"][-1] >= 3)
    if got != want:
        err("文字摘要內部名單二不符（%d／%d）" % (sum(got.values()), sum(want.values())))
    for ar in dec:
        if "〔%s〕" % ar not in sec2:
            err("文字摘要內部名單二缺少 %s" % ar)

    # 9 主日都沒來，但小排或家聚會有來（三份報表模式）
    if aux:
        want = bm("近12週主日都沒來，但年底小排或家聚會有紀錄：%d人，其中年底來2週以上%d人") % (len(NS), sum(1 for z in NS if z["gq"] + z["hq"] >= 2))
        if want not in txt:
            err("摘要缺少：" + want)
        sec3 = internal.split("【四、近12週主日都沒來，但小排或家聚會有來】")
        if len(sec3) != 2:
            err("文字摘要內部名單缺少第三段")
        else:
            got = [l for l in sec3[1].splitlines() if l.startswith("　")]
            exp3 = [bm("　%s（%s｜近一年主日%d次｜近一年小排%d週、家聚會%d週｜年底小排%d週、家聚會%d週）") % (
                z["name"], z["grp"] or "無群組", z["sy"], z["gy"], z["hy"], z["gq"], z["hq"]) for z in NS]
            if got != exp3:
                err("文字摘要內部名單三不符（%d／%d）" % (len(got), len(exp3)))
            if "共%d人" % len(NS) not in sec3[1]:
                err("文字摘要內部名單三總人數不符")
        cells = list(wb["近一年概況"].iter_rows(values_only=True))
        i4 = next((n for n, r in enumerate(cells) if str(r[0] or "").startswith("五、近12週主日都沒來")), None)
        if i4 is None:
            err("近一年概況缺少第四段")
        else:
            got = [tuple(r[:8]) for r in cells[i4 + 3:i4 + 3 + len(NS)]]
            exp4 = [(z["area"], z["name"], z["grp"] or None, z["sy"], z["gy"], z["hy"], z["gq"], z["hq"]) for z in NS]
            got = [(a, b, c or None, d, e, f, g, h) for a, b, c, d, e, f, g, h in got]
            if got != exp4:
                err("近一年概況第四段名單不符")

    # 9-1 增加的區（2026-11-13 新增）：多來3次以上的人、增減拆解
    def dlt(p):
        return p["sc"][-1] - p["sc"][0]
    grow = [u for u in ADULT + ["兒童"] if avg(U[u], 0) and (avg(U[u], K - 1) - avg(U[u], 0)) / avg(U[u], 0) >= 0.05]
    got = Counter(re.findall(r"^　(.+)（[^（｜]+｜[^｜]+｜[^｜]*多來", sec_up, re.M))
    want = Counter(p["name"] for u in grow for p in U[u] if dlt(p) >= 3)
    if got != want:
        err("文字摘要內部名單三（增加的區）不符（%d／%d）" % (sum(got.values()), sum(want.values())))
    for u in grow:
        if "〔%s〕" % u not in sec_up:
            err("文字摘要內部名單三缺少 %s" % u)
    cells_x = list(wb["近一年概況"].iter_rows(values_only=True))
    j4 = next((n for n, r in enumerate(cells_x) if str(r[0] or "").startswith("四、人數增加")), None)
    j5 = next((n for n, r in enumerate(cells_x) if str(r[0] or "").startswith("五、")), len(cells_x))
    if j4 is None:
        err("近一年概況缺少第四段（增加的區）")
    else:
        got = Counter((r[0], r[1]) for r in cells_x[j4 + 3:j5] if r[0] and r[1])
        want = Counter((u, p["name"]) for u in grow for p in U[u] if dlt(p) >= 3)
        if got != want:
            err("近一年概況第四段（增加的區）名單不符（%d／%d）" % (sum(got.values()), sum(want.values())))

    def split_head(ps, up):
        big = [p for p in ps if (dlt(p) >= 3 if up else -dlt(p) >= 3)]
        tot = sum(dlt(p) for p in ps) / 12
        bs = sum(dlt(p) for p in big) / 12
        of = sum(1 for p in big if (p["ls"][-1] if up else p["ls"][0]) in ("穩定", "普通"))
        word, who = ("多來", "年底已常來") if up else ("少來", "年初常來")
        return bm("%s3次以上的%d人合計%+.1f人（其中%s%d人），其餘%+.1f人" % (word, len(big), bs, who, of, tot - bs))
    dec_ = [ar for ar in ADULT if avg(U[ar], 0) and (avg(U[ar], K - 1) - avg(U[ar], 0)) / avg(U[ar], 0) <= -0.10]
    for u in grow:
        if split_head(U[u], True) not in sec_up:
            err("文字摘要內部名單三 %s 增減拆解不符" % u)
    for u in dec_:
        if split_head(U[u], False) not in sec2:
            err("文字摘要內部名單二 %s 增減拆解不符" % u)

    # 10 重點摘要（2026-11-13 新增）：名單、區別判斷與增減來源
    hp = os.path.join(a.outdir, "近一年重點摘要_%s.txt" % tag)
    if not os.path.exists(hp):
        err("找不到重點摘要檔")
    else:
        hl = open(hp, encoding="utf-8").read()
        if "年初" in hl or "年底" in hl:
            err("重點摘要仍有「年初」「年底」字樣")
        ao = {x: i for i, x in enumerate(AREAS)}
        allp = U["全會所"] + U["兒童"]
        kf = lambda p: (ao[p["area"]], p["name"])

        def names_in(label):
            if "〔%s" % label not in hl:
                return None
            blk = hl.split("〔%s" % label, 1)[1].split("\n〔", 1)[0].split("\n【", 1)[0]
            return [re.match(r"^　(.+)（[^（｜]+｜", l).group(1) for l in blk.splitlines() if l.startswith("　") and l != "　無"]
        want = {
            bm("淡出：年初常來，年底都沒來"): [p["name"] for p in sorted([p for p in allp if p["t"] == "淡出"], key=kf)],
            "常來但近一個月都沒來": [p["name"] for p in sorted([p for p in allp if p["t"] == "常來但近一個月沒來"], key=kf)],
            "下滑，且近一個月都沒來": [p["name"] for p in sorted([p for p in allp if p["t"] == "下滑" and p["r4"] == 0], key=kf)],
            "本區間第二季以後才開始出現，曾常來但近一個月沒來": [p["name"] for p in sorted([p for p in allp if p["new"] and p["ns"] == "曾常來但近一個月沒來"], key=kf)],
        }
        if aux:
            want["共"] = [z["name"] for z in NS if z["gq"] + z["hq"] >= 2]
        for label, names in want.items():
            got = names_in(label)
            if got != names:
                err("重點摘要〔%s〕名單不符：應為 %s；實際 %s" % (label, "、".join(names) or "無", "、".join(got or []) or "無"))
        sec3 = hl.split("【三、")[1].split("【四、")[0] if "【三、" in hl else ""
        advice = {"一直不常來": "本來就不常來的人來得更少", "下滑": "年底變得不常來的人", "淡出": "年底都沒來的人",
                  "常來但近一個月沒來": "近一個月都沒來的人", "本區間第二季以後才開始出現": "之後沒有留下來", "留住": "固定來的人都來得少",
                  "回升": "回升的人又來得較少"}
        for ar in ADULT:
            a0, a1 = avg(U[ar], 0), avg(U[ar], K - 1)
            line = next((l for l in sec3.splitlines() if l.startswith("・%s：平均每週" % ar)), None)
            if a0 and (a1 - a0) / a0 <= -0.10:
                c = contrib(U[ar])
                t = min(c, key=lambda k: c[k])
                if not line or bm(advice.get(t, t)) not in line:
                    err("重點摘要第三點 %s 的減少來源判斷不符（應為「%s」）" % (ar, t))
            elif line:
                err("重點摘要第三點不應列出 %s" % ar)
        sec2 = hl.split("【二、")[1].split("【三、")[0] if "【二、" in hl else ""
        for u in ADULT + ["兒童"]:
            a0, a1 = avg(U[u], 0), avg(U[u], K - 1)
            grow = bool(a0) and (a1 - a0) / a0 >= 0.05
            if grow != ("・%s：平均每週" % u in sec2):
                err("重點摘要第二點 %s %s" % (u, "應列出" if grow else "不應列出"))
        for sec in ("【一、", "【二、", "【三、", "【四、"):
            part = hl.split(sec)[1].split("\n【")[0] if sec in hl else ""
            leak = [p["name"] for p in P if len(p["name"]) >= 2 and p["name"] in part]
            if leak:
                err("重點摘要%s段不應出現姓名：%s" % (sec, "、".join(leak[:3])))

    sec4 = txt.split("【四、各區增減來自哪些人】")[1].split("【五、")[0] if "【四、各區增減來自哪些人】" in txt else ""
    for u in ADULT + ["兒童", "青職（跨區）", "大專（跨區）"]:
        tot = sum(contrib(U[u]).values())
        if "%s（%+.1f）：" % (u, tot) not in sec4:
            err("摘要第四點缺少或不符：%s" % u)

    if errors:
        print("【回測失敗】")
        for e in errors[:40]:
            print("・" + e)
        sys.exit(1)
    print("【回測通過】近一年主日%d次、%d段；分析對象%d人（兒童%d）；本區間第二季以後才開始出現%d人；訊號%d項" % (
        W, K, len(P), len(U["兒童"]), sum(1 for p in P if p["new"]), len(sig)))


if __name__ == "__main__":
    main()
