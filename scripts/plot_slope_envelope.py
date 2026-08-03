# -*- coding: utf-8 -*-
"""オンスペック中の「1時間あたりの調整幅」の実績の幅を、計画と同じ起点・同じ3時間で引いて比較する。"""
import pickle, sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams['font.family'] = 'IPAGothic'
plt.rcParams['axes.unicode_minus'] = False
SURF, INK, INK2, MUTED = '#fcfcfb', '#0b0b0b', '#52514e', '#a8a7a1'
BLUE, ORANGE, VIOLET = '#2a78d6', '#eb6834', '#4a3aa7'

df, su, ev, smry, hourly = pickle.load(open(sys.argv[1], 'rb'))
L, Q = df['FC1402'], df['FQY1530']
adj = ev[(ev['局面'] == 'オンスペック中の稼働調整') & (~ev['トリップ・緊急操作の疑い'])]


def hourly_steps(series, floor, link_load=False):
    """調整イベントの中で、変化方向と同じ向きに動いた1時間あたりの変化幅。
    link_load=True のときは、同じ時間に負荷(FC1402)も同方向に動いているものだけを採る
    （MEG側の払出流量の一時的な振れを実績から除くため）。"""
    out = []
    for _, r in adj.iterrows():
        d = series.loc[r['開始']:r['終了']].diff()
        sg = np.sign(r['変化量_TpH'])
        m = (np.sign(d) == sg) & (d.abs() > floor)
        if link_load:
            dl = L.loc[r['開始']:r['終了']].diff()
            m = m & (np.sign(dl) == sg) & (dl.abs() > 0.05)
        for v in d[m.fillna(False)].values:
            out.append((abs(v), '増' if sg > 0 else '減'))
    return pd.DataFrame(out, columns=['幅', '向き'])


SPEC = [
    dict(name='EOフィード  FC1402', ser=L, floor=0.05, link=False, y0=7.5, tgt=12.999,
         plan=1.8331, steps=[('step1 ロードダウン', 0.6233), ('step2 ロードダウン', 0.4767),
                             ('テスト後戻し ロードアップ', 1.8331)], ymax=15.5),
    dict(name='MEG生産量  FQY1530', ser=Q, floor=0.50, link=True, y0=7.625, tgt=15.833,
         plan=2.7361, steps=[('step1 ロードダウン', 0.9583), ('step2 ロードダウン', 0.6833),
                             ('テスト後戻し ロードアップ', 2.7361)], ymax=21.5),
]
for s in SPEC:
    s['H'] = hourly_steps(s['ser'], s['floor'], s['link'])
    up = s['H'][s['H']['向き'] == '増']['幅']
    s['up'] = up
    s['q'] = {k: float(up.quantile(v)) for k, v in
              [('中央値', .5), ('p75', .75), ('p90', .90), ('p99', .99)]}
    s['q']['最大'] = float(up.max())

fig = plt.figure(figsize=(13.6, 10.4), facecolor=SURF)
gs = fig.add_gridspec(2, 2, height_ratios=[1.25, 1.0], hspace=0.36, wspace=0.18,
                      left=0.062, right=0.985, top=0.905, bottom=0.075)

t = np.linspace(0, 3, 61)
for i, s in enumerate(SPEC):
    ax = fig.add_subplot(gs[0, i])
    ax.set_facecolor(SURF)
    lo, hi = s['q']['中央値'], s['q']['最大']
    ax.fill_between(t, s['y0'] + lo * t, s['y0'] + hi * t, color=MUTED, alpha=0.28, lw=0,
                    label='実績の幅（1時間あたりの調整幅 中央値〜最大）')
    ax.plot(t, s['y0'] + hi * t, lw=2.2, color=INK2,
            label='実績の最大  %.2f T/H per h' % hi)
    for k in ['p99', 'p90', '中央値']:
        v = s['q'][k]
        ax.plot(t, s['y0'] + v * t, lw=1.4, color=MUTED, ls=(0, (5, 3)))
        ax.annotate('%s %.2f' % (k, v), xy=(3, s['y0'] + v * 3), xytext=(4, -3),
                    textcoords='offset points', fontsize=8.5, color=INK2)
    ax.annotate('最大 %.2f' % hi, xy=(3, s['y0'] + hi * 3), xytext=(4, -3),
                textcoords='offset points', fontsize=8.5, color=INK)
    ax.plot([0, 1, 2, 3], [s['y0'] + s['plan'] * k for k in range(4)], lw=2.8, color=ORANGE,
            marker='o', ms=7, zorder=5,
            label='今回計画  %.2f T/H per h' % s['plan'])
    ax.axhline(s['tgt'], color=ORANGE, lw=0.9, ls=(0, (4, 3)), zorder=1)
    ax.annotate('計画の到達値 %.1f T/H' % s['tgt'], xy=(0.05, s['tgt']), xytext=(0, 5),
                textcoords='offset points', ha='left', fontsize=9, color=ORANGE)
    ax.set_title('%s（起点 %.2f T/H から3時間）' % (['①', '②'][i] + ' ' + s['name'], s['y0']),
                 fontsize=11.5, color=INK, loc='left', pad=9)
    ax.set_xlabel('増量開始からの経過時間 [h]', fontsize=9.5, color=INK2)
    ax.set_ylabel('流量 [T/H]', fontsize=9.5, color=INK2)
    ax.set_xlim(0, 3.45)
    ax.set_ylim(s['y0'] - 0.5, s['ymax'])
    ax.legend(frameon=False, fontsize=9, loc='upper left')
    for sp in ('top', 'right'):
        ax.spines[sp].set_visible(False)
    for sp in ('left', 'bottom'):
        ax.spines[sp].set_color(MUTED)
    ax.tick_params(colors=INK2, labelsize=9)
    ax.grid(axis='y', color=MUTED, alpha=0.35, lw=0.6)
    ax.set_axisbelow(True)

for i, s in enumerate(SPEC):
    ax = fig.add_subplot(gs[1, i])
    ax.set_facecolor(SURF)
    up = s['up']
    bins = np.linspace(0, s['q']['最大'] * 1.05, 40)
    ax.hist(up, bins=bins, color=MUTED, alpha=0.75, lw=0)
    for k in ['中央値', 'p90', 'p99', '最大']:
        v = s['q'][k]
        ax.axvline(v, color=INK2, lw=1.1, ls=(0, (5, 3)))
        ax.annotate('%s %.2f' % (k, v), xy=(v, 0.97), xycoords=('data', 'axes fraction'),
                    xytext=(3, 0), textcoords='offset points', fontsize=8.5, color=INK2, va='top')
    for j, (nm, v) in enumerate(s['steps']):
        c = ORANGE if 'ロードアップ' in nm else VIOLET
        ax.axvline(v, color=c, lw=2.4, zorder=5)
        ax.annotate('%s\n%.2f' % (nm.split(' ')[0], v), xy=(v, 0.30 + 0.16 * (j % 2)),
                    xycoords=('data', 'axes fraction'), xytext=(4, 0),
                    textcoords='offset points', fontsize=8.5, color=c, va='center')
    pct = (up <= s['plan']).mean() * 100
    ax.set_title('%s %s：1時間あたりの調整幅の実績分布（増量 N=%d）\n'
                 '　　計画のロードアップ %.2f は実績の %.1f パーセンタイル・実績最大 %.2f'
                 % (['③', '④'][i], s['name'], len(up), s['plan'], pct, s['q']['最大']),
                 fontsize=10.5, color=INK, loc='left', pad=9)
    ax.set_xlabel('1時間あたりの変化幅 [T/H per h]', fontsize=9.5, color=INK2)
    ax.set_ylabel('時間数', fontsize=9.5, color=INK2)
    ax.set_yscale('log')
    for sp in ('top', 'right'):
        ax.spines[sp].set_visible(False)
    for sp in ('left', 'bottom'):
        ax.spines[sp].set_color(MUTED)
    ax.tick_params(colors=INK2, labelsize=9)
    ax.grid(axis='y', color=MUTED, alpha=0.3, lw=0.6)
    ax.set_axisbelow(True)

fig.suptitle('オンスペック中の「1時間あたりの調整幅」の実績の幅と、今回のテスト計画の位置',
             fontsize=14.5, color=INK, x=0.062, ha='left', y=0.968)
fig.savefig(sys.argv[2], dpi=170, facecolor=SURF)
print('saved', sys.argv[2])
for s in SPEC:
    print(s['name'], {k: round(v, 2) for k, v in s['q'].items()},
          '計画', s['plan'], 'pct %.1f' % ((s['up'] <= s['plan']).mean() * 100))
