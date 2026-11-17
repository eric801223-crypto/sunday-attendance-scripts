# -*- coding: utf-8 -*-
"""參考優先關心名單 互動式 HTML（以近12週短期情形為主｜原型）

讀取已通過回測的「主日出席分析_YYYYMMDD.xlsx」，不重新分類，依類別（同卡片：類別名稱、白話說明、數字條件）
呈現每個人近12週的出席情形。可另給最近一次「近一年出席分析_YYYYMMDD.xlsx」，在點開個人時附一行近一年參考
（註明該份的製表日期；兩份同一天時，另核對12週紀錄與近一年紀錄最後12週相同）。
產出前自我核對：各區各類別人數、＊人數、青職與大專人數與 Excel 相同，相符才顯示「【HTML完成】」。

用法：python3 -I care_html.py 主日出席分析_YYYYMMDD.xlsx 輸出.html [--annual 近一年出席分析_YYYYMMDD.xlsx]
"""
import json
import re
import sys
from collections import defaultdict, Counter

import openpyxl

AREAS = ['1小區', '2小區', '3小區', '4小區', '5小區', '6小區', '7小區', '青年區', '兒童專項']
TYPES = ['留住', '常來但近一個月沒來', '下滑', '淡出', '回升', '年中開始出現', '一直不常來']
RANKS = ['出席穩定者本週缺席', '出席穩定者連續缺席兩週以上', '出席普通者近期缺席', '出席普通者連續缺席三週以上',
         '總出席低，但最近一個月來了一半以上', '總出席低，但最近一個月有來1次',
         '總出席低，原本隔週來，但最近一個月都沒來', '總出席低，原本很少來，且最近一個月都沒來',
         '過去不時出現，但近12週都沒來']


def stop(msg):
    print('【停止】' + msg)
    sys.exit(1)


def need(wb, name):
    if name not in wb.sheetnames:
        stop(f'找不到工作表「{name}」，Excel 格式可能已改變。')
    return wb[name]


def header_ok(ws, row, expect, label):
    got = [c.value for c in ws[row]][:len(expect)]
    if got != expect:
        stop(f'{label}的欄位與預期不同：{got}')


def main(care_x, out, ann_x=None):
    cw = openpyxl.load_workbook(care_x, data_only=True)
    aw = openpyxl.load_workbook(ann_x, data_only=True) if ann_x else None

    # ── 基本資訊 ──
    cmade = re.search(r'製表日期：(\S+)', need(cw, '重點摘要').cell(2, 1).value).group(1)
    period = need(cw, '重點摘要').cell(3, 1).value.replace('統計期間：', '')
    rdef = {}
    for r in need(cw, '分類概況').iter_rows(values_only=True):
        m = re.match(r'第(\d)順位：(.+?)｜(.+?)（(.+)）$', str(r[0] or ''))
        if m:
            rdef[int(m.group(1))] = {'name': m.group(2), 'plain': m.group(3), 'cond': m.group(4)}
    if sorted(rdef) != list(range(1, 10)):
        stop('分類概況的類別說明（第1～9順位）讀取不完整。')
    for k, v in rdef.items():
        if v['name'] != RANKS[k - 1]:
            stop(f'第{k}類別名稱「{v["name"]}」與預期不同。')

    # ── 近一年參考（選用）──
    ann, amade, same_day, seg_label, wseg = defaultdict(list), None, False, [], []
    if aw:
        ov = need(aw, '近一年概況')
        amade = re.search(r'製表日期：(\S+)', ov.cell(2, 1).value).group(1)
        same_day = amade == cmade
        segs = re.findall(r'第(\d)段（(\d{4})/(\d{2})/\d{2}～(\d{4})/(\d{2})/\d{2}）', ov.cell(4, 1).value)
        seg_label = [f'{int(a)}～{int(b)}月' for _, _, a, _, b in segs]
        wp = need(aw, '個人近一年軌跡（內部）')
        header_ok(wp, 4, ['所屬區', '兒童', '姓名', '群組', '第1段', '第2段', '第3段', '第4段', '等級軌跡', '近4週', '類型',
                          '近一年出席', '最近出席日期', '小排或家聚會各段週數', '近一年出席紀錄（左舊右新）'], '個人近一年軌跡')
        for r in wp.iter_rows(min_row=5, values_only=True):
            if r[0] is None:
                continue
            ann[(r[0], r[2])].append({'type': r[10], 's': list(r[4:8]), 'year': r[11],
                                      'dots': ''.join('1' if c == '●' else '0' for c in r[14])})

    def join(area, name, dots12):
        cand = ann.get((area, name), [])
        if not aw:
            return None, ''
        if not cand:
            return None, '不在近一年分析對象（全年主日0次）'
        if same_day and dots12 is not None:
            hit = [c for c in cand if c['dots'][-12:] == dots12]
            if len(hit) == 1:
                return hit[0], ''
            if len(cand) == 1:
                stop(f'{area} {name}：兩份 Excel 同一天製表，但12週紀錄與近一年紀錄最後12週不同。')
            return None, '同區同名，無法比對'
        if len(cand) == 1:
            return cand[0], ''
        return None, '同區同名，無法比對'

    # ── 群組（12次出席矩陣）──
    mx = need(cw, '12次出席矩陣')
    header_ok(mx, 1, ['所屬區', '姓名', '群組'], '12次出席矩陣')
    grp = {}
    dates = [c.value for c in mx[1]][3:15]
    for r in mx.iter_rows(min_row=2, values_only=True):
        if r[0]:
            grp[(r[0], r[1])] = r[2]

    # ── 關心名單 ──
    ws = need(cw, '全區看望總表')
    header_ok(ws, 1, ['順位', '類別', '所屬區', '姓名', '最近出席日期', '連續缺席週數', '近四週出席', '12週出席',
                      '近12週以前一年內出席', '最近12次出席紀錄（左舊右新）', '小排近一季', '家聚會近一季',
                      '小排近一個月', '家聚會近一個月', '標＊'], '全區看望總表')
    care = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        if r[0] is None:
            continue
        d12 = ''.join('1' if c == '●' else '0' for c in r[9])
        a, why = join(r[2], r[3], d12)
        care.append({'list': 'care', 'ord': i, 'rank': r[0], 'cat': r[1], 'area': r[2], 'name': r[3],
                     'group': grp.get((r[2], r[3]), ''), 'last': r[4], 'k': r[5], 'n4w': r[6], 'n12': r[7],
                     'prev': r[8], 'd12': d12, 'gq': r[10], 'hq': r[11], 'gm': r[12], 'hm': r[13],
                     'star': r[14] == '＊', 'a': a, 'why': why})

    # ── 接近穩定（內部）──
    wn = need(cw, '接近穩定（內部）')
    header_ok(wn, 4, ['所屬區', '姓名', '群組', '12週出席', '近四週出席', '連續缺席週數', '目前情形',
                      '最近12次出席紀錄（左舊右新）'], '接近穩定')
    near = []
    for i, r in enumerate(wn.iter_rows(min_row=5, values_only=True)):
        if r[0] is None:
            continue
        d12 = ''.join('1' if c == '●' else '0' for c in r[7])
        a, why = join(r[0], r[1], d12)
        near.append({'list': 'near', 'ord': i, 'area': r[0], 'name': r[1], 'group': r[2], 'n12': r[3], 'n4w': r[4],
                     'k': r[5], 'state': r[6], 'd12': d12, 'a': a, 'why': why})

    # ── 主日沒來但有小排家聚會（內部）──
    contact = []
    if '主日沒來但有小排家聚會（內部）' in cw.sheetnames:
        wc = cw['主日沒來但有小排家聚會（內部）']
        header_ok(wc, 3, ['所屬區', '姓名', '群組', '小排近一季', '家聚會近一季', '小排近一個月', '家聚會近一個月'], '主日沒來但有小排家聚會')
        for i, r in enumerate(wc.iter_rows(min_row=4, values_only=True)):
            if r[0] is None:
                continue
            a, why = join(r[0], r[1], '0' * 12)
            if a is None and why == '同區同名，無法比對':
                a, why = join(r[0], r[1], None)
            contact.append({'list': 'contact', 'ord': i, 'area': r[0], 'name': r[1], 'group': r[2],
                            'gq': r[3], 'hq': r[4], 'gm': r[5], 'hm': r[6], 'a': a, 'why': why})

    # ── 自我核對 ──
    errs = []
    sm = need(cw, '分類概況')
    header_ok(sm, 4, ['分區', '有效名單', '第1順位', '第2順位', '第3順位', '第4順位', '第5順位', '第6順位', '第7順位',
                      '第8順位', '第9順位', '出席穩定（本週有來）', '出席普通（本週有來）', '標＊人數'], '分類概況')
    cnt = Counter((p['area'], p['rank']) for p in care)
    stars = Counter(p['area'] for p in care if p['star'])
    for r in sm.iter_rows(min_row=5, values_only=True):
        if r[0] not in AREAS:
            continue
        calc = [cnt[(r[0], k)] for k in range(1, 10)]
        if calc != list(r[2:11]):
            errs.append(f'{r[0]} 各順位人數：重算 {calc}，Excel {list(r[2:11])}')
        if stars[r[0]] != r[13]:
            errs.append(f'{r[0]} ＊人數：重算 {stars[r[0]]}，Excel {r[13]}')
    for g, sheet in (('青職', '青職分類概況'), ('大學', '大專分類概況')):
        ref = {}
        for r in need(cw, sheet).iter_rows(min_row=2, values_only=True):
            m = re.match(r'第(\d)順位', str(r[0] or ''))
            if m:
                ref[int(m.group(1))] = r[1]
        calc = Counter(p['rank'] for p in care if p['group'] == g)
        if any(calc[k] != ref.get(k) for k in range(1, 10)):
            errs.append(f'{sheet}：重算 {[calc[k] for k in range(1, 10)]}，Excel {[ref.get(k) for k in range(1, 10)]}')
    if errs:
        stop('HTML 自我核對不符：\n' + '\n'.join(errs))

    allp = care + near + contact
    for p in allp:
        if p['a'] and p['a']['type'] not in TYPES:
            stop(f'{p["name"]} 的近一年類型「{p["a"]["type"]}」不在預期類型中。')
    unmatched = Counter((p['list'], p['why']) for p in allp if p['a'] is None)

    data = {'made': cmade, 'period': period, 'amade': amade, 'seg': seg_label, 'dates': dates,
            'areas': AREAS, 'rdef': rdef, 'care': care, 'near': near, 'contact': contact}
    html = open(__file__.replace('care_html.py', 'care_html_template.html'), encoding='utf-8').read()
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html.replace('/*DATA*/', json.dumps(data, ensure_ascii=False)))
    um = '未提供' if not aw else '、'.join(f'{ {"care": "關心名單", "near": "接近穩定", "contact": "主日沒來但有小排家聚會"}[l]}{w} {n}人'
                  for (l, w), n in unmatched.items()) or '無'
    print(f'【HTML完成】{out}（關心名單{len(care)}人、接近穩定{len(near)}人、主日沒來但有小排家聚會{len(contact)}人；'
          f'各區各類別、＊、青職、大專人數與 Excel 相同；近一年參考：{amade or "未提供"}，對不上者：{um}）')


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('care'); ap.add_argument('out'); ap.add_argument('--annual')
    a = ap.parse_args()
    main(a.care, a.out, a.annual)
