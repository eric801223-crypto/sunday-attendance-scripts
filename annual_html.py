# -*- coding: utf-8 -*-
"""近一年出席分析 互動式 HTML（原型）

讀取已通過回測的「近一年出席分析_YYYYMMDD.xlsx」，不重新分類，只把同一份數字換成互動圖表。
產出前先自我核對：由個人資料重算的每週人數、前後季對照、細項、增減拆解，必須與 Excel 相同，
相同才寫出 HTML 並顯示「【HTML完成】」；不同就以「【停止】」結束。

同一次也做關心名單時，另給該次「主日出席分析_YYYYMMDD.xlsx」（--care），在名單中加上每人近12週在關心名單的類別
（寫成「近12週：類別名稱」；兩份同一天時核對12週紀錄與近一年紀錄最後12週相同）。只做近一年分析時不加 --care。

用法：python3 -I annual_html.py 近一年出席分析_YYYYMMDD.xlsx 輸出.html [--care 主日出席分析_YYYYMMDD.xlsx]
"""
import json
import re
import sys
from collections import defaultdict

import openpyxl

TYPES = ['留住', '常來但近一個月沒來', '下滑', '淡出', '回升', '年中開始出現', '一直不常來']
REGIONS = ['1小區', '2小區', '3小區', '4小區', '5小區', '6小區', '7小區', '青年區']
UNITS = ['全會所'] + REGIONS + ['兒童', '青職', '大專']
LEVELS = ['穩定', '普通', '不常來', '沒來']


def stop(msg):
    print('【停止】' + msg)
    sys.exit(1)


def level(n):
    return '穩定' if n >= 9 else '普通' if n >= 6 else '不常來' if n >= 1 else '沒來'


def sheet(wb, name):
    if name not in wb.sheetnames:
        stop(f'找不到工作表「{name}」，Excel 格式可能已改變。')
    return wb[name]


def in_unit(p, u):
    child = p['child']
    if u == '全會所':
        return not child
    if u == '兒童':
        return child
    if u == '青職':
        return p['group'] == '青職' and not child
    if u == '大專':
        return p['group'] == '大學' and not child
    return p['region'] == u and not child


def main(xlsx, out, care_x=None):
    wb = openpyxl.load_workbook(xlsx, data_only=True)

    # ── 基本資訊 ──
    ov = sheet(wb, '近一年概況')
    head = [ov.cell(r, 1).value or '' for r in range(1, 8)]
    made = re.search(r'製表日期：(\S+)', head[1]).group(1)
    period = head[2].replace('近一年期間：', '')
    segs = re.findall(r'第(\d)段（(\d{4})/(\d{2})/\d{2}～(\d{4})/(\d{2})/\d{2}）', head[3])
    if len(segs) != 4:
        stop('近一年概況第4列的分段格式與預期不同。')
    seg_label = [f'{int(a)}～{int(b)}月' for _, _, a, _, b in segs]
    mode = head[6].replace('分析模式：', '')

    # ── 每週趨勢 ──
    wt = sheet(wb, '每週趨勢')
    hdr = [c.value for c in wt[2]]
    expect = ['主日日期', '段別', '全會所（不含訪客、兒童）', '4週平均'] + REGIONS + ['兒童（另計）', '訪客等臨時點名（另計）']
    if hdr[:14] != expect:
        stop('每週趨勢的欄位與預期不同。')
    weeks = []
    for r in wt.iter_rows(min_row=3, values_only=True):
        if r[0] is None:
            break
        weeks.append({'date': str(r[0])[:10], 'seg': r[1], 'total': r[2],
                      'reg': dict(zip(REGIONS, r[4:12])), 'child': r[12], 'visitor': r[13]})
    hl = []
    for r in wt.iter_rows(min_row=3 + len(weeks), values_only=True):
        if r[0] in ('最高', '最低'):
            hl.append({'kind': r[0], 'date': str(r[1])[:10], 'n': r[2]})

    # ── 個人軌跡 ──
    wp = sheet(wb, '個人近一年軌跡（內部）')
    if [c.value for c in wp[4]][:15] != ['所屬區', '兒童', '姓名', '群組', '第1段', '第2段', '第3段', '第4段',
                                         '等級軌跡', '近4週', '類型', '近一年出席', '最近出席日期',
                                         '小排或家聚會各段週數', '近一年出席紀錄（左舊右新）']:
        stop('個人近一年軌跡的欄位與預期不同。')
    people = []
    for r in wp.iter_rows(min_row=5, values_only=True):
        if r[0] is None:
            continue
        dots = r[14]
        if len(dots) != len(weeks) or set(dots) - {'●', '○'}:
            stop(f'{r[2]} 的出席紀錄長度或符號與每週趨勢不符。')
        gh = [int(x) if x.strip().isdigit() else None for x in (r[13] or '').split('／')] if r[13] else None
        people.append({'region': r[0], 'child': r[1] == '是' or r[0] == '兒童專項', 'name': r[2], 'group': r[3],
                       's': list(r[4:8]), 'n4': r[9], 'type': r[10], 'year': r[11], 'last': r[12],
                       'gh': gh, 'dots': ''.join('1' if c == '●' else '0' for c in dots)})
    for p in people:
        if p['type'] not in TYPES:
            stop(f'{p["name"]} 的類型「{p["type"]}」不在預期類型中。')

    # ── 年中開始出現 ──
    wn = sheet(wb, '年中開始出現（內部）')
    newc = {}
    for r in wn.iter_rows(min_row=5, values_only=True):
        if r[0] is None:
            continue
        newc[(r[0], r[2])] = {'first': str(r[4])[:10], 'n12': r[5], 'status': r[7], 'gh': r[8]}

    # ── Excel 上的比對基準 ──
    wq = sheet(wb, '前後季對照')
    q_sum, q_mat, cur = {}, {}, None
    for r in wq.iter_rows(min_row=5, values_only=True):
        if r[0] in [u if u in ('全會所', '兒童') or u in REGIONS else f'{u}（跨區）' for u in UNITS] and r[1] is not None and not str(r[0]).startswith('【'):
            q_sum[str(r[0]).replace('（跨區）', '')] = list(r[1:10])
        if r[0] and str(r[0]).startswith('【'):
            cur = str(r[0]).strip('【】').replace('（跨區）', '')
            q_mat[cur] = []
        elif cur and r[0] and re.match(r'\d+～\d+月(穩定|普通|不常來|沒來)$', str(r[0])):
            q_mat[cur].append(list(r[1:5]))
    dec = {}
    for row in range(1, ov.max_row + 1):
        if ov.cell(row, 1).value == '區' and ov.cell(row, 5).value == '留住':
            for rr in range(row + 1, row + 20):
                v = [ov.cell(rr, c).value for c in range(1, 12)]
                if v[0] is None:
                    break
                dec[str(v[0]).replace('（跨區）', '')] = {'a': v[1], 'b': v[2], 'd': v[3], 'by': dict(zip(TYPES, v[4:11]))}
            break
    wx = sheet(wb, '主日×小排家聚會')
    cross, cur = {}, None
    for r in wx.iter_rows(values_only=True):
        if r[0] and str(r[0]).startswith('【'):
            cur = str(r[0]).strip('【】').replace('（跨區）', '')
            cross[cur] = []
        elif cur and r[0] and re.search(r'主日(穩定|普通|不常來|沒來)$', str(r[0])):
            cross[cur].append([r[1], r[2]])

    # ── 自我核對 ──
    errs = []
    for u in UNITS:
        ps = [p for p in people if in_unit(p, u)]
        # 前後季對照摘要
        common0 = [p for p in ps if level(p['s'][0]) in ('穩定', '普通') and p['type'] != '年中開始出現']
        calc = [len(common0),
                sum(1 for p in common0 if level(p['s'][3]) in ('穩定', '普通')),
                sum(1 for p in ps if p['type'] == '常來但近一個月沒來'),
                sum(1 for p in ps if p['type'] == '下滑'),
                sum(1 for p in ps if p['type'] == '淡出'),
                sum(1 for p in ps if p['type'] == '回升'),
                sum(1 for p in ps if p['type'] == '年中開始出現'),
                sum(1 for p in ps if p['type'] == '一直不常來')]
        xl = q_sum.get(u)
        if xl is None:
            errs.append(f'前後季對照找不到 {u}')
        else:
            xl_cmp = [xl[0], xl[1], xl[2], xl[3], xl[4], xl[6], xl[7], xl[8]]
            if calc != xl_cmp:
                errs.append(f'前後季對照 {u}：重算 {calc}，Excel {xl_cmp}')
        # 細項
        m = [[0] * 4 for _ in range(4)]
        for p in ps:
            m[LEVELS.index(level(p['s'][0]))][LEVELS.index(level(p['s'][3]))] += 1
        if q_mat.get(u) != m:
            errs.append(f'細項 {u}：重算 {m}，Excel {q_mat.get(u)}')
        # 每週人數
        wk = [sum(int(p['dots'][i]) for p in ps) for i in range(len(weeks))]
        ref = ([w['total'] for w in weeks] if u == '全會所' else [w['child'] for w in weeks] if u == '兒童'
               else [w['reg'][u] for w in weeks] if u in REGIONS else None)
        if ref is not None and wk != ref:
            errs.append(f'每週人數 {u} 與每週趨勢不同')
        # 增減拆解
        if u in dec:
            for t in TYPES:
                v = sum(p['s'][3] - p['s'][0] for p in ps if p['type'] == t) / 12
                if abs(v - (dec[u]['by'][t] or 0)) > 0.05 + 1e-9:  # Excel 為四捨五入到一位小數
                    errs.append(f'增減拆解 {u} {t}：重算 {v:.3f}，Excel {dec[u]["by"][t]}')
        # 主日×小排
        if mode.endswith('家聚會') and u in cross:
            cm = [[0, 0] for _ in range(4)]
            for p in ps:
                if p['gh'] is None or p['gh'][3] is None:
                    continue
                cm[LEVELS.index(level(p['s'][3]))][0 if p['gh'][3] > 0 else 1] += 1
            if cm != cross[u]:
                errs.append(f'主日×小排 {u}：重算 {cm}，Excel {cross[u]}')
    if errs:
        stop('HTML 自我核對不符：\n' + '\n'.join(errs))

    for p in people:
        k = (p['region'], p['name'])
        if p['type'] == '年中開始出現' and k in newc:
            p['new'] = newc[k]

    # ── 目前關心名單（選用）──
    cmade, rdef, cinfo = None, {}, {}
    if care_x:
        cw = openpyxl.load_workbook(care_x, data_only=True)
        cmade = re.search(r'製表日期：(\S+)', sheet(cw, '重點摘要').cell(2, 1).value).group(1)
        for r in sheet(cw, '分類概況').iter_rows(values_only=True):
            m = re.match(r'第(\d)順位：(.+?)｜(.+?)（(.+)）$', str(r[0] or ''))
            if m:
                rdef[int(m.group(1))] = {'name': m.group(2), 'plain': m.group(3), 'cond': m.group(4)}
        if sorted(rdef) != list(range(1, 10)):
            stop('關心名單 Excel 的類別說明（第1～9順位）讀取不完整。')
        ws = sheet(cw, '全區看望總表')
        if [c.value for c in ws[1]][:10] != ['順位', '類別', '所屬區', '姓名', '最近出席日期', '連續缺席週數', '近四週出席',
                                             '12週出席', '近12週以前一年內出席', '最近12次出席紀錄（左舊右新）']:
            stop('全區看望總表的欄位與預期不同。')
        byname = defaultdict(list)
        for i, p in enumerate(people):
            byname[(p['region'], p['name'])].append(i)
        miss = 0
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] is None:
                continue
            d12 = ''.join('1' if c == '●' else '0' for c in r[9])
            cand = byname.get((r[2], r[3]), [])
            if cmade == made:
                hit = [i for i in cand if people[i]['dots'][-12:] == d12]
                if len(cand) == 1 and not hit:
                    stop(f'{r[2]} {r[3]}：兩份 Excel 同一天製表，但12週紀錄與近一年紀錄最後12週不同。')
            else:
                hit = cand
            if len(hit) == 1:
                people[hit[0]]['care'] = {'rank': r[0], 'star': r[14] == '＊', 'k': r[5]}
            else:
                miss += 1
        print(f'關心名單對應：{sum(1 for p in people if "care" in p)}人；對不上（全年主日0次或同區同名）{miss}人')

    data = {'cmade': cmade, 'rdef': rdef, 'made': made, 'period': period, 'seg': seg_label, 'mode': mode, 'weeks': weeks, 'hl': hl,
            'people': people, 'units': UNITS, 'types': TYPES, 'levels': LEVELS,
            'dec': dec, 'cross': cross}
    html = TEMPLATE.replace('/*DATA*/', json.dumps(data, ensure_ascii=False))
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'【HTML完成】{out}（{len(people)}人、{len(weeks)}週；自我核對 {len(UNITS)} 個單位全部相同）')


TEMPLATE = open(__file__.replace('annual_html.py', 'annual_html_template.html'), encoding='utf-8').read()

if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx'); ap.add_argument('out'); ap.add_argument('--care')
    a = ap.parse_args()
    main(a.xlsx, a.out, a.care)
