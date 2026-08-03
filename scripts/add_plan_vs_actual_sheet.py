# -*- coding: utf-8 -*-
"""計画（テスト後戻しのロード調整表）と、オンスペック中の増量実績の重ね書きシートを追加する。"""
import pickle, sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

df, su, ev, smry, hourly = pickle.load(open(sys.argv[1], 'rb'))
L, Q = df['FC1402'], df['FQY1530']
on = (L > 0.5) & (Q > 0.5)
wb = load_workbook(sys.argv[2])
ws = wb.create_sheet('計画と実績_重ね書き')
B = Font(bold=True)

ws['A1'] = '今回のテスト後戻し計画と、オンスペック中の増量実績の比較（FC1402 R-1401 EO FEED）'
ws['A1'].font = B
rows = [
    ['計画 テスト後戻し(切替前) 8/7', 'EOフィード 7.5 → 13.0 T/H', '+5.5 T/H を 3 時間', '1.83 T/H per h', '1時間あたり +1.83 T/H'],
    ['計画 テスト後戻し(LERA切替後) 8/6', 'EOフィード 7.5 → 13.0 T/H', '+5.5 T/H を 4 時間', '1.37 T/H per h', '1時間あたり +1.37 T/H'],
    ['実績 オンスペック中の増量 3 T/H超（26件）', '', '中央値 +5.3 T/H', '中央値 0.40 T/H per h', 'p90 0.66 / 最大 1.02'],
    ['実績 段階的な増量の最速（2025/8/1）', '11.60 → 17.70 T/H', '+6.1 T/H を 5 時間', '1.02 T/H per h', '1時間の最大 +2.11 T/H'],
    ['実績 最速（2022/5/7）', '9.27 → 17.48 T/H', '+8.2 T/H を 2 時間', '4.11 T/H per h', '1時間の最大 +4.53 T/H'],
]
ws['A3'] = 'ケース'; ws['B3'] = '負荷'; ws['C3'] = '変化量と所要時間'; ws['D3'] = '平均速度'; ws['E3'] = '1時間あたり'
for c in 'ABCDE':
    ws['%s3' % c].font = B
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        ws.cell(row=4 + i, column=1 + j, value=v)

r0 = 11
ws.cell(row=r0 - 1, column=1, value='重ね書き用データ（増量開始を0時間、開始時からの増分 [T/H]）').font = B
past = [('2020-03-09 10:00', 4), ('2020-07-27 10:00', 4), ('2021-01-06 10:00', 4),
        ('2022-05-07 10:00', 4), ('2020-08-30 07:00', 10), ('2021-01-28 07:00', 9),
        ('2025-08-01 09:00', 8)]
plan3 = [0, 1.833111, 3.666222, 5.499333]
plan4 = [0, 1.374833, 2.749667, 4.124500, 5.499333]
series = [('計画 切替前 3h', plan3), ('計画 LERA切替後 4h', plan4)]
for t0, hrs in past:
    seg = L.loc[t0:pd.Timestamp(t0) + pd.Timedelta(hours=hrs)]
    series.append((pd.Timestamp(t0).strftime('%Y/%-m/%-d 実績'), list(seg.values - seg.iloc[0])))
nmax = max(len(v) for _, v in series)
ws.cell(row=r0, column=1, value='経過時間_h').font = B
for i in range(nmax):
    ws.cell(row=r0 + 1 + i, column=1, value=i)
for j, (name, vals) in enumerate(series):
    ws.cell(row=r0, column=2 + j, value=name).font = B
    for i, v in enumerate(vals):
        c = ws.cell(row=r0 + 1 + i, column=2 + j, value=float(v))
        c.number_format = '0.00'

ch = LineChart()
ch.title = '増量開始からの FC1402 増分 [T/H]'
ch.y_axis.title = 'FC1402 の増分 [T/H]'
ch.x_axis.title = '経過時間 [h]'
ch.height, ch.width = 10, 20
data = Reference(ws, min_col=2, max_col=1 + len(series), min_row=r0, max_row=r0 + nmax)
cats = Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + nmax)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ws.add_chart(ch, '%s%d' % (get_column_letter(len(series) + 4), r0))
for i, w in enumerate([34, 26, 22, 20, 22] + [16] * len(series), start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
wb.save(sys.argv[2])
print('updated', sys.argv[2])
