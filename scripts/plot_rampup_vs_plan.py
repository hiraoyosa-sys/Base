# -*- coding: utf-8 -*-
"""オンスペック中の増量実績と、テスト後戻し計画（EGロード調整）の重ね書き。"""
import pickle, sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager

plt.rcParams['font.family'] = 'IPAGothic'
plt.rcParams['axes.unicode_minus'] = False

SURF = '#fcfcfb'
INK = '#0b0b0b'
INK2 = '#52514e'
MUTED = '#a8a7a1'
BLUE, ORANGE, VIOLET = '#2a78d6', '#eb6834', '#4a3aa7'

df, su, ev, smry, hourly = pickle.load(open(sys.argv[1], 'rb'))
L, Q = df['FC1402'], df['FQY1530']
on = (L > 0.5) & (Q > 0.5)

# オンスペック中の増量イベント（3 T/H以上）
onev = ev[(ev['局面'] == 'オンスペック中の稼働調整') & (ev['変化量_TpH'] >= 3.0)].copy()
# 3時間で+5.5 T/H以上を達成した「特に速い」実績（窓全体がオンスペック）
w = 3
okm = (on.rolling(w + 1).sum() == w + 1).fillna(False)
d3 = L.diff(w)[okm].dropna()
fast_hours = d3[d3 >= 5.5].index
fast_days = sorted({t.normalize() for t in fast_hours})
# S/U当日（FQY1530が前バッチの残流量で立っているだけ）の 2020-06-25 は除外
fast_days = [d for d in fast_days if d != pd.Timestamp('2020-06-25')]

# 計画（添付2 ロード調整表）
plan3 = [7.5, 9.333111, 11.166222, 12.999333]                        # テスト後戻し(切替前) 3h

fig = plt.figure(figsize=(13.5, 12.4), facecolor=SURF)
gs = fig.add_gridspec(3, 1, height_ratios=[1.0, 1.1, 1.1], hspace=0.32,
                      left=0.065, right=0.985, top=0.93, bottom=0.055)

# ---------------- 上：どの時期か ----------------
ax = fig.add_subplot(gs[0])
ax.set_facecolor(SURF)
s = L.resample('6h').mean()
ax.plot(s.index, s.values, lw=0.9, color=MUTED, zorder=1)
lbl1 = True
for _, r in onev.iterrows():
    ax.plot([r['開始'], r['終了']], [r['開始負荷_TpH'], r['終了負荷_TpH']], color=ORANGE, lw=2.6,
            solid_capstyle='round', zorder=3,
            label='オンスペック中の増量 +3 T/H以上（26件）' if lbl1 else None)
    lbl1 = False
lbl2 = True
for d in fast_days:                      # 3時間窓で+5.5 T/H以上を達成した日
    day = d3.loc[d:d + pd.Timedelta(hours=23)]
    tpk = day.idxmax()
    seg = L.loc[tpk - pd.Timedelta(hours=w):tpk]
    ax.plot(seg.index, seg.values, color=VIOLET, lw=3.6, solid_capstyle='round', zorder=5,
            label='3時間で+5.5 T/H以上（今回計画と同等以上の速さ）' if lbl2 else None)
    lbl2 = False
    ax.annotate(d.strftime('%Y/%-m/%-d'), xy=(tpk, seg.max()), xytext=(0, 8),
                textcoords='offset points', ha='center', fontsize=8.5, color=VIOLET)
ax.set_ylabel('FC1402  R-1401 EO FEED [T/H]', fontsize=10, color=INK2)
ax.set_title('① オンスペック中に大きく増量した時期（2020/1〜2026/5、細線＝6時間平均の負荷）',
             fontsize=12, color=INK, loc='left', pad=10)
ax.legend(frameon=False, fontsize=9.5, loc='lower left', ncol=2)
ax.set_ylim(0, 23)
for sp in ('top', 'right'):
    ax.spines[sp].set_visible(False)
for sp in ('left', 'bottom'):
    ax.spines[sp].set_color(MUTED)
ax.tick_params(colors=INK2, labelsize=9)
ax.grid(axis='y', color=MUTED, alpha=0.35, lw=0.6)
ax.set_axisbelow(True)

# ---------------- 下：計画と実績の重ね書き ----------------
ax2 = fig.add_subplot(gs[1])
ax2.set_facecolor(SURF)
past = [('2020-03-09 10:00', 4), ('2020-07-27 10:00', 4), ('2021-01-06 10:00', 4),
        ('2022-05-07 10:00', 4), ('2020-08-30 07:00', 10), ('2021-01-28 07:00', 9)]
first = True
for t0, hrs in past:
    seg = L.loc[t0:pd.Timestamp(t0) + pd.Timedelta(hours=hrs)]
    ax2.plot(range(len(seg)), seg.values, lw=1.8, color=MUTED, zorder=2,
             label='過去のオンスペック中の増量' if first else None)
    first = False
    ax2.annotate(pd.Timestamp(t0).strftime('%-m/%-d'), xy=(len(seg) - 1, seg.iloc[-1]),
                 xytext=(5, -3), textcoords='offset points', fontsize=8.5, color=INK2)
seg = L.loc['2025-08-01 09:00':'2025-08-01 17:00']
ax2.plot(range(len(seg)), seg.values, lw=2.6, color=BLUE, zorder=4,
         label='2025/8/1  直近の段階増量  11.6 → 17.7 T/H を5時間（1.02 T/H per h）')
ax2.plot(range(len(plan3)), np.array(plan3), lw=2.8, color=ORANGE, zorder=5, marker='o', ms=7,
         label='計画 テスト後戻し(切替前)  7.5 → 13.0 T/H を3時間（1.83 T/H per h）')
ax2.axhline(13.0, color=ORANGE, lw=0.9, ls=(0, (4, 3)), zorder=1)
ax2.annotate('計画の到達値 13.0 T/H', xy=(11.4, 13.0), xytext=(0, 5), textcoords='offset points',
             ha='right', fontsize=9, color=ORANGE)
ax2.set_xlabel('増量開始からの経過時間 [h]', fontsize=10, color=INK2)
ax2.set_ylabel('FC1402  R-1401 EO FEED [T/H]', fontsize=10, color=INK2)
ax2.set_title('② EOフィード（FC1402）：計画と、オンスペック中の増量実績の重ね書き（縦軸は流量そのもの）', fontsize=12, color=INK, loc='left', pad=10)
ax2.legend(frameon=False, fontsize=9.5, loc='lower right')
ax2.set_xlim(-0.2, 11.5)
ax2.set_ylim(6, 21.5)
for sp in ('top', 'right'):
    ax2.spines[sp].set_visible(False)
for sp in ('left', 'bottom'):
    ax2.spines[sp].set_color(MUTED)
ax2.tick_params(colors=INK2, labelsize=9)
ax2.grid(axis='y', color=MUTED, alpha=0.35, lw=0.6)
ax2.set_axisbelow(True)

# ---------------- 下：MEG（FQY1530）の上昇 ----------------
Qc = Q.where(Q > 5)                       # タンクヤード切替でQ=0に落ちる時間は除外
ax3 = fig.add_subplot(gs[2])
ax3.set_facecolor(SURF)
meg_past = [('2020-01-30 12:00', 18), ('2020-03-24 14:00', 12), ('2020-08-30 07:00', 13),
            ('2021-01-28 07:00', 11), ('2021-08-02 15:00', 10), ('2024-10-15 13:00', 11),
            ('2025-05-14 08:00', 12)]
planQ3 = [7.625, 10.361111, 13.097222, 15.833333]
first = True
for t0, hrs in meg_past:
    seg = Qc.loc[t0:pd.Timestamp(t0) + pd.Timedelta(hours=hrs)].ffill()
    ax3.plot(range(len(seg)), seg.values, lw=1.8, color=MUTED, zorder=2,
             label='過去のオンスペック中の増量' if first else None)
    first = False
    ax3.annotate(pd.Timestamp(t0).strftime('%-m/%-d'), xy=(len(seg) - 1, seg.iloc[-1]),
                 xytext=(5, -3), textcoords='offset points', fontsize=8.5, color=INK2)
seg = Qc.loc['2025-08-01 09:00':'2025-08-01 19:00'].ffill()
ax3.plot(range(len(seg)), seg.values, lw=2.6, color=BLUE, zorder=4,
         label='2025/8/1  実績の最速  12.7 → 21.9 T/H を6時間（1.52 T/H per h）')
ax3.plot(range(len(planQ3)), np.array(planQ3), lw=2.8, color=ORANGE, zorder=5, marker='o', ms=7,
         label='計画 テスト後戻し(切替前)  7.6 → 15.8 T/H を3時間（2.74 T/H per h）')
ax3.axhline(15.83, color=ORANGE, lw=0.9, ls=(0, (4, 3)), zorder=1)
ax3.annotate('計画の到達値 15.8 T/H', xy=(18.6, 15.83), xytext=(0, 5), textcoords='offset points',
             ha='right', fontsize=9, color=ORANGE)
ax3.set_xlabel('増量開始からの経過時間 [h]', fontsize=10, color=INK2)
ax3.set_ylabel('FQY1530  MEG TO TANKYARD [T/H]', fontsize=10, color=INK2)
ax3.set_title('③ MEG生産量（FQY1530）：計画と、オンスペック中の増量実績の重ね書き（縦軸は流量そのもの）',
              fontsize=12, color=INK, loc='left', pad=10)
ax3.legend(frameon=False, fontsize=9.5, loc='lower right')
ax3.set_xlim(-0.3, 19)
ax3.set_ylim(6, 38)
for sp in ('top', 'right'):
    ax3.spines[sp].set_visible(False)
for sp in ('left', 'bottom'):
    ax3.spines[sp].set_color(MUTED)
ax3.tick_params(colors=INK2, labelsize=9)
ax3.grid(axis='y', color=MUTED, alpha=0.35, lw=0.6)
ax3.set_axisbelow(True)

fig.suptitle('EGロード（FC1402）とMEG生産量（FQY1530）の増量速度：今回のテスト後戻し計画は実績の範囲内か',
             fontsize=14.5, color=INK, x=0.065, ha='left', y=0.975)
fig.savefig(sys.argv[2], dpi=170, facecolor=SURF)
print('saved', sys.argv[2])
