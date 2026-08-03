# -*- coding: utf-8 -*-
"""
FQY1530(MEG TO TANKYARD) / FC1402(R-1401 EO FEED) の1時間トレンドから
稼働調整（負荷変更）の速度実績を、MEGオンスペック時 と S/U直後（未オンスペック）時 に分けて集計する。

入力 : MEG検討ベース.xlsm の「トレンドデータ」シート
        A列=日時(Excelシリアル), B列=EOG.FQY1530.PV [T/H], C列=EOG.FC1402.PV [T/H]
出力 : outputs/FQY1530_FC1402_稼働調整速度.xlsx

判定の考え方
  運転中          : FC1402 > 0.5 T/H
  オンスペック     : FQY1530 > 0.5 T/H （MEGをタンクヤードへ払い出している＝製品出荷中）
  稼働調整イベント : FC1402 の1時間差が 0.05 T/H を超える時間の連なり（2時間までの中断は同一イベント）
                    かつ 区間の正味変化が 0.30 T/H 以上
  速度            : 区間の正味変化量 ÷ 区間の所要時間 [T/H per h]
"""
import sys, zipfile, io, pickle, datetime as dt
import xml.etree.ElementTree as ET
import numpy as np
import pandas as pd

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RUN_TH, ONS_TH = 0.5, 0.5      # T/H
NOISE, GAP, MINSTEP = 0.05, 2, 0.30
TRIP_1H = 3.0                  # 1時間で3 T/H超の変化はトリップ・緊急操作として調整実績から除外


def read_trend(xlsm_path, sheet='xl/worksheets/sheet6.xml'):
    z = zipfile.ZipFile(xlsm_path)
    shared = []
    for _, el in ET.iterparse(io.BytesIO(z.read('xl/sharedStrings.xml'))):
        if el.tag == NS + 'si':
            shared.append(''.join(t.text or '' for t in el.iter(NS + 't')))
            el.clear()
    rows = []
    for _, el in ET.iterparse(z.open(sheet), events=('end',)):
        if el.tag != NS + 'row':
            continue
        r = int(el.get('r'))
        if r >= 11:
            d = {}
            for c in el.findall(NS + 'c'):
                col = ''.join(ch for ch in c.get('r') if ch.isalpha())
                v = c.find(NS + 'v')
                val = v.text if v is not None else None
                if c.get('t') == 's' and val is not None:
                    val = shared[int(val)]
                d[col] = val
            rows.append((d.get('A'), d.get('B'), d.get('C')))
        el.clear()
    df = pd.DataFrame(rows, columns=['serial', 'FQY1530', 'FC1402'])
    for c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce')   # PIの "No Good Data" 等は欠測に落とす
    df = df.dropna(subset=['serial'])
    df['ts'] = (pd.Timestamp('1899-12-30') + pd.to_timedelta(df['serial'], unit='D')).dt.round('h')
    return df.set_index('ts')[['FQY1530', 'FC1402']].sort_index()


def campaigns(L, min_h=24):
    run = (L > RUN_TH).fillna(False)
    g = (run != run.shift()).cumsum()
    out = []
    for _, idx in L.groupby(g).groups.items():
        if run.loc[idx].iloc[0] and len(idx) >= min_h:
            out.append((idx[0], idx[-1]))
    return out


def su_table(L, Q, camps):
    ons = (Q > ONS_TH).rolling(3).sum() == 3
    rows = []
    for s, e in camps:
        if s == L.index[0]:
            continue          # データ先頭は運転途中からの切り出し＝S/Uではない
        seg = L.loc[s:e]
        hit = ons.loc[s:e]
        hit = hit[hit].index
        t_on = hit[0] - pd.Timedelta(hours=2) if len(hit) else pd.NaT
        rec = {'EO投入': s, '停止': e, '運転時間_h': (e - s) / pd.Timedelta(hours=1),
               'オンスペック時刻': t_on,
               'EO投入からオンスペックまで_h': np.nan, 'オンスペック時の負荷_TpH': np.nan,
               '初期ランプ所要_h': np.nan, '初期ランプ変化量_TpH': np.nan,
               '初期ランプ速度_TpH_per_h': np.nan,
               '初期ランプ中の最大1h上昇_TpH_per_h': np.nan,
               'ホールド時間_h': np.nan, 'ホールド中の正味変化_TpH': np.nan,
               'ホールド中の最大1h変化_TpH_per_h': np.nan,
               'キャンペーン最大負荷_TpH': seg.max()}
        if pd.notna(t_on):
            L_on = L.loc[max(t_on - pd.Timedelta(hours=2), s):t_on].median()
            rec['EO投入からオンスペックまで_h'] = (t_on - s) / pd.Timedelta(hours=1)
            rec['オンスペック時の負荷_TpH'] = L_on
            reach = seg.loc[:t_on]
            hit2 = reach[reach >= 0.95 * L_on]
            if len(hit2):
                t_r = hit2.index[0]
                h = max((t_r - s) / pd.Timedelta(hours=1), 1)
                rec['初期ランプ所要_h'] = (t_r - s) / pd.Timedelta(hours=1)
                rec['初期ランプ変化量_TpH'] = reach.loc[t_r] - reach.iloc[0]
                rec['初期ランプ速度_TpH_per_h'] = (reach.loc[t_r] - reach.iloc[0]) / h
                rec['初期ランプ中の最大1h上昇_TpH_per_h'] = reach.loc[:t_r].diff().max()
                hold = L.loc[t_r:t_on]
                rec['ホールド時間_h'] = (t_on - t_r) / pd.Timedelta(hours=1)
                rec['ホールド中の正味変化_TpH'] = hold.iloc[-1] - hold.iloc[0]
                rec['ホールド中の最大1h変化_TpH_per_h'] = hold.diff().abs().max()
        rows.append(rec)
    return pd.DataFrame(rows)


def detect_events(L):
    dL = L.diff()
    act = ((dL.abs() > NOISE) & dL.notna()).values
    idx, vals, n = L.index, L.values, len(L)
    out, i = [], 0
    while i < n:
        if not act[i]:
            i += 1
            continue
        last, k = i, i
        while k < n:
            if act[k]:
                last, k = k, k + 1
            elif k - last <= GAP:
                k += 1
            else:
                break
        st, en = idx[max(i - 1, 0)], idx[last]
        seg = L.loc[st:en]
        if seg.notna().all() and len(seg) >= 2:
            d = seg.iloc[-1] - seg.iloc[0]
            dur = (en - st) / pd.Timedelta(hours=1)
            if abs(d) >= MINSTEP:
                out.append({'開始': st, '終了': en, '所要_h': dur,
                            '開始負荷_TpH': seg.iloc[0], '終了負荷_TpH': seg.iloc[-1],
                            '変化量_TpH': d, '速度_TpH_per_h': d / dur,
                            '最大1h変化_TpH_per_h': seg.diff().abs().max() * np.sign(d)})
        i = last + 1
    return pd.DataFrame(out)


def classify(ev, L, Q, su):
    ramp_end = {}
    for _, r in su.iterrows():
        if pd.notna(r['初期ランプ所要_h']):
            ramp_end[r['EO投入']] = r['EO投入'] + pd.Timedelta(hours=float(r['初期ランプ所要_h']))

    def f(r):
        if r['終了負荷_TpH'] < 1.0 and r['開始負荷_TpH'] >= 3.0:
            return 'SD・トリップ（0まで降下）'
        if r['開始負荷_TpH'] < 1.0 and r['終了負荷_TpH'] >= 3.0:
            return 'S/U初期ランプ（0→初期負荷・未オンスペック）'
        q = Q.loc[r['開始']:r['終了']]
        onspec = (q > ONS_TH).mean() >= 0.5 if len(q) else False
        if onspec:
            return 'オンスペック中の稼働調整'
        # 未オンスペックで運転中：S/U後オンスペック前か、それ以外か
        for s, te in ramp_end.items():
            if s <= r['開始'] <= te + pd.Timedelta(hours=200):
                return '未オンスペック（S/U後・初期負荷ホールド期）'
        return '未オンスペック（オンスペック未達キャンペーン）'
    ev = ev.copy()
    ev['局面'] = ev.apply(f, axis=1)
    ev['向き'] = np.where(ev['変化量_TpH'] > 0, '増', '減')
    ev['トリップ・緊急操作の疑い'] = ev['最大1h変化_TpH_per_h'].abs() > TRIP_1H
    d = ev['変化量_TpH'].abs()
    ev['変化量区分'] = np.select([d < 1, d < 3], ['小（0.3〜1 T/H）', '中（1〜3 T/H）'], '大（3 T/H超）')
    ev['速度_絶対値'] = ev['速度_TpH_per_h'].abs()
    ev['相対速度_pct_per_h'] = ev['速度_絶対値'] / ev['開始負荷_TpH'].clip(lower=0.5) * 100
    return ev


def summarize(ev, su):
    def q(x, p):
        x = np.asarray(x, dtype=float)
        x = x[~np.isnan(x)]
        return float(np.percentile(x, p)) if len(x) else np.nan

    def m(x):
        x = np.asarray(x, dtype=float)
        x = x[~np.isnan(x)]
        return float(np.median(x)) if len(x) else np.nan

    def mx_(x):
        x = np.asarray(x, dtype=float)
        x = x[~np.isnan(x)]
        return float(np.max(x)) if len(x) else np.nan
    rows = []

    def add(name, n, rate, step, dur, mx):
        rows.append({'局面': name, '件数': n,
                     '速度_中央値_TpH_per_h': m(rate),
                     '速度_p25': q(rate, 25), '速度_p75': q(rate, 75), '速度_p90': q(rate, 90),
                     '速度_最大': mx_(rate),
                     '1回の変化量_中央値_TpH': m(step),
                     '所要時間_中央値_h': m(dur),
                     '最大1h変化_中央値_TpH_per_h': m(mx),
                     '最大1h変化_最大_TpH_per_h': mx_(mx)})

    s = su.dropna(subset=['初期ランプ速度_TpH_per_h'])
    add('S/U初期ランプ（EO投入→初期負荷・未オンスペック）', len(s),
        s['初期ランプ速度_TpH_per_h'].abs().values, s['初期ランプ変化量_TpH'].abs().values,
        s['初期ランプ所要_h'].values, s['初期ランプ中の最大1h上昇_TpH_per_h'].values)
    h = su.dropna(subset=['ホールド中の最大1h変化_TpH_per_h'])
    add('初期負荷ホールド（初期負荷到達→オンスペック・未オンスペック）', len(h),
        (h['ホールド中の正味変化_TpH'].abs() / h['ホールド時間_h'].clip(lower=1)).values,
        h['ホールド中の正味変化_TpH'].abs().values, h['ホールド時間_h'].values,
        h['ホールド中の最大1h変化_TpH_per_h'].values)

    adj = ev[(ev['局面'] == 'オンスペック中の稼働調整') & (~ev['トリップ・緊急操作の疑い'])]
    for k in ['小（0.3〜1 T/H）', '中（1〜3 T/H）', '大（3 T/H超）']:
        g = adj[adj['変化量区分'] == k]
        add('オンスペック中の稼働調整　' + k, len(g), g['速度_絶対値'].values,
            g['変化量_TpH'].abs().values, g['所要_h'].values, g['最大1h変化_TpH_per_h'].abs().values)
    add('オンスペック中の稼働調整　合計', len(adj), adj['速度_絶対値'].values,
        adj['変化量_TpH'].abs().values, adj['所要_h'].values, adj['最大1h変化_TpH_per_h'].abs().values)
    t = ev[ev['局面'] == 'SD・トリップ（0まで降下）']
    add('（参考）SD・トリップの降下', len(t), t['速度_絶対値'].values,
        t['変化量_TpH'].abs().values, t['所要_h'].values, t['最大1h変化_TpH_per_h'].abs().values)
    return pd.DataFrame(rows)


def main(xlsm, out_xlsx):
    df = read_trend(xlsm)
    L, Q = df['FC1402'], df['FQY1530']
    camps = campaigns(L)
    su = su_table(L, Q, camps)
    ev = classify(detect_events(L), L, Q, su)
    smry = summarize(ev, su)
    hourly = pd.DataFrame({
        '局面': ['運転中・オンスペック', '運転中・未オンスペック', '停止'],
    })
    onm = (L > RUN_TH) & (Q > ONS_TH)
    offm = (L > RUN_TH) & (Q <= ONS_TH)
    sdm = (L <= RUN_TH)
    recs = []
    for name, m in [('運転中・オンスペック', onm), ('運転中・未オンスペック', offm), ('停止', sdm)]:
        d = L.diff()[m & m.shift(fill_value=False)].dropna().abs()
        recs.append({'局面': name, '時間数': len(d), '中央値': d.median(),
                     'p75': d.quantile(.75), 'p90': d.quantile(.90), 'p95': d.quantile(.95),
                     'p99': d.quantile(.99), '最大': d.max(),
                     '静止（1h変化0.05未満）の割合': (d < 0.05).mean()})
    hourly = pd.DataFrame(recs)
    return df, su, ev, smry, hourly


if __name__ == '__main__':
    xlsm = sys.argv[1]
    out = sys.argv[2]
    df, su, ev, smry, hourly = main(xlsm, out)
    pickle.dump((df, su, ev, smry, hourly), open('/tmp/claude-0/-home-user-Base/f5434811-769d-5ff1-bb6c-1bc71c19ed94/scratchpad/result.pkl', 'wb'))
    print(smry.round(3).to_string(index=False))
    print()
    print(hourly.round(4).to_string(index=False))
