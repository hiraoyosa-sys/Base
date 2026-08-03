# -*- coding: utf-8 -*-
"""fqy1530_fc1402_ramp_analysis.py の解析結果を1ブックにまとめる。"""
import sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

sys.path.insert(0, '/home/user/Base/scripts')
from fqy1530_fc1402_ramp_analysis import main as analyze

HDR = Font(bold=True)


def put(ws, r, c, v, bold=False):
    cell = ws.cell(row=r, column=c, value=v)
    if bold:
        cell.font = HDR
    return cell


def write_df(ws, df, r0, c0=1, datefmt='yyyy/mm/dd hh:mm', numfmt='0.000'):
    for j, col in enumerate(df.columns):
        put(ws, r0, c0 + j, str(col), bold=True)
    for i, (_, row) in enumerate(df.iterrows()):
        for j, col in enumerate(df.columns):
            v = row[col]
            if isinstance(v, pd.Timestamp):
                cell = ws.cell(row=r0 + 1 + i, column=c0 + j, value=v.to_pydatetime())
                cell.number_format = datefmt
            elif isinstance(v, (np.floating, float)):
                cell = ws.cell(row=r0 + 1 + i, column=c0 + j,
                               value=None if pd.isna(v) else float(v))
                cell.number_format = numfmt
            elif isinstance(v, (np.integer, int, np.bool_, bool)):
                ws.cell(row=r0 + 1 + i, column=c0 + j, value=int(v) if not isinstance(v, (bool, np.bool_)) else bool(v))
            else:
                ws.cell(row=r0 + 1 + i, column=c0 + j, value=None if pd.isna(v) else str(v))
    return r0 + len(df) + 1


def autow(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build(xlsm, out):
    df, su, ev, smry, hourly = analyze(xlsm, out)
    wb = Workbook()

    # ---------------- サマリ ----------------
    ws = wb.active
    ws.title = 'サマリ'
    r = 1
    put(ws, r, 1, 'FQY1530 / FC1402 による稼働調整速度の実績整理', bold=True); r += 2
    for line in [
        ['対象データ', 'MEG検討ベース.xlsm 「トレンドデータ」シート（1時間値）'],
        ['期間', '%s 〜 %s（%d時間）' % (df.index[0].strftime('%Y/%m/%d %H:%M'),
                                    df.index[-1].strftime('%Y/%m/%d %H:%M'), len(df))],
        ['負荷の指標', 'EOG.FC1402.PV  R-1401 EO FEED [T/H]（EG向EO消費）'],
        ['オンスペックの指標', 'EOG.FQY1530.PV  MEG TO TANKYARD [T/H]（0.5 T/H超＝タンクヤード払出中）'],
        ['運転中の判定', 'FC1402 が 0.5 T/H 超'],
        ['調整イベントの抽出', 'FC1402 の1時間差が 0.05 T/H 超の連なり（2時間までの中断は同一イベント）、'
                        'かつ区間の正味変化 0.30 T/H 以上'],
        ['速度の定義', '区間の正味変化量 ÷ 区間の所要時間 [T/H per h]'],
        ['除外', '1時間で 3 T/H を超える変化を含むイベントは、トリップ・緊急操作として調整実績から除外'],
        ['欠測', 'PI の No Good Data For Calculation は欠測扱い（FQY1530 517時間、FC1402 1,043時間）'],
    ]:
        put(ws, r, 1, line[0], bold=True); put(ws, r, 2, line[1]); r += 1
    r += 1
    put(ws, r, 1, '局面別の調整速度 [T/H per h]', bold=True); r += 1
    r = write_df(ws, smry, r) + 1
    put(ws, r, 1, '1時間ごとの変化量 |ΔFC1402| の分布 [T/H per h]（イベント抽出をせず全時間で集計）', bold=True); r += 1
    r = write_df(ws, hourly, r, numfmt='0.000') + 1
    put(ws, r, 1, '読み取り', bold=True); r += 1
    for t in [
        'S/U直後（MEG未オンスペック）は、EO投入から初期負荷まで中央値 1.03 T/H per h（p75 2.04、最大 2.99）で一気に上げている。'
        '1時間だけを見ると中央値 3.1 T/H per h、最大 7.3 T/H per h まで振っている。',
        '初期負荷に達したあとオンスペックになるまでは、負荷をほぼ動かしていない（正味変化の中央値 0.40 T/H、'
        '1時間変化の最大でも中央値 0.37 T/H per h）。EO投入からオンスペックまでは中央値 21.5 時間（12〜42時間）。',
        'オンスペック中の稼働調整は中央値 0.13 T/H per h。変化量が大きいほど速度も上げており、'
        '3 T/H超の増減産で中央値 0.42 T/H per h（p90 0.66、最大 1.65）、1時間の最大でも中央値 1.11 T/H per h。',
        '同じ大きさの負荷変更で比べると、S/U初期ランプはオンスペック中の大調整のおよそ2.5倍の速度、'
        '1時間の最大値では約3倍で動かしている。',
        '注意：FQY1530 はタンクヤードへの払出流量であり、EG系（R-1401）以外からのMEG合流を含む可能性がある'
        '（FQY1530/FC1402 の比の中央値 1.30、量論上限 62/44=1.41 を超える時間が 42%）。'
        'オンスペックかどうかの判定（0か0超か）には影響しないが、MEG生産量そのものの評価に使う場合は要確認。',
    ]:
        put(ws, r, 1, t); r += 1
    autow(ws, [26, 22] + [14] * 10)
    ws.freeze_panes = 'A2'

    # ---------------- S/U一覧 ----------------
    ws2 = wb.create_sheet('S_U一覧')
    put(ws2, 1, 1, 'S/U（EO投入）ごとの立上げ実績。オンスペック時刻は FQY1530 が 0.5 T/H 超で3時間続いた最初の時刻。', bold=True)
    su2 = su.copy()
    su2['備考'] = np.where(su2['EO投入からオンスペックまで_h'] > 72, 'オンスペックまで72時間超（低負荷運転を挟んだ特殊ケース）',
                         np.where(su2['EO投入からオンスペックまで_h'].isna(), 'オンスペックに至らずに停止', ''))
    write_df(ws2, su2, 3)
    autow(ws2, [19, 19, 12, 19, 20, 18, 14, 16, 20, 24, 12, 16, 22, 18, 18, 40])
    ws2.freeze_panes = 'A4'

    # ---------------- 調整イベント一覧 ----------------
    ws3 = wb.create_sheet('調整イベント一覧')
    put(ws3, 1, 1, 'FC1402 の連続した変化を1件＝1イベントとして抽出したもの。', bold=True)
    ev2 = ev.sort_values('開始').reset_index(drop=True)
    write_df(ws3, ev2, 3)
    autow(ws3, [19, 19, 9, 13, 13, 12, 16, 20, 34, 6, 22, 18, 12, 18])
    ws3.freeze_panes = 'A4'

    # ---------------- 代表S/Uトレンド ----------------
    ws4 = wb.create_sheet('代表S_Uトレンド')
    put(ws4, 1, 1, '直近3回のS/U前後の1時間値。FC1402＝負荷、FQY1530＝MEGタンクヤード払出。', bold=True)
    windows = [('2024-06-26 12:00', '2024-06-29 12:00'),
               ('2025-02-16 12:00', '2025-02-19 12:00'),
               ('2025-07-22 00:00', '2025-07-25 00:00')]
    col = 1
    for a, b in windows:
        seg = df.loc[a:b]
        put(ws4, 3, col, 'S/U %s' % a[:10], bold=True)
        put(ws4, 4, col, '経過時間_h', bold=True)
        put(ws4, 4, col + 1, 'FC1402_TpH', bold=True)
        put(ws4, 4, col + 2, 'FQY1530_TpH', bold=True)
        for i, (t, row) in enumerate(seg.iterrows()):
            ws4.cell(row=5 + i, column=col, value=float((t - seg.index[0]) / pd.Timedelta(hours=1)))
            for k, c in enumerate(['FC1402', 'FQY1530']):
                cell = ws4.cell(row=5 + i, column=col + 1 + k,
                                value=None if pd.isna(row[c]) else float(row[c]))
                cell.number_format = '0.00'
        ch = LineChart()
        ch.title = 'S/U %s' % a[:10]
        ch.y_axis.title = 'T/H'
        ch.x_axis.title = 'EO投入前からの経過時間 [h]'
        ch.height, ch.width = 7, 12
        data = Reference(ws4, min_col=col + 1, max_col=col + 2, min_row=4, max_row=4 + len(seg))
        cats = Reference(ws4, min_col=col, min_row=5, max_row=4 + len(seg))
        ch.add_data(data, titles_from_data=True)
        ch.set_categories(cats)
        ws4.add_chart(ch, '%s3' % get_column_letter(col + 4))
        col += 13
    wb.save(out)
    print('saved:', out)


if __name__ == '__main__':
    build(sys.argv[1], sys.argv[2])
