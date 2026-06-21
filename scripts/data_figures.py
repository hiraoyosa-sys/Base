# -*- coding: utf-8 -*-
"""データ図（淡いMC配色）：APHA経時／UVスペクトル模式／UV450早期判定。配布クリーン。"""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

fp = fm.FontProperties(fname="/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf")
def F(sz, bold=False):
    f = fm.FontProperties(fname="/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", size=sz)
    if bold: f.set_weight("bold")
    return f
BLUE = "#005BAB"; DBLUE = "#003F7E"; AMBER = "#E0A000"; GRAY = "#888888"; GREEN = "#2E7D32"
OUTDIR = os.path.join(os.path.dirname(__file__), "..", "outputs")
os.makedirs(OUTDIR, exist_ok=True)

def save(fig, name):
    p = os.path.join(OUTDIR, name)
    fig.savefig(p, bbox_inches="tight", facecolor="white", dpi=200); plt.close(fig)
    print("saved:", os.path.normpath(p))

# ── 1) APHA経時（タンク保管） ──
fig, ax = plt.subplots(figsize=(6.4, 3.2))
d = np.array([0, 23, 26, 43, 51, 60, 70])
tank = np.array([3, 5, 10, 15, 22, 28, 31])
ax.plot(d, tank, "-o", color=BLUE, lw=2.2, ms=5, label="タンク品（窒素封入）")
ax.axhline(30, 0, 1, color=AMBER, lw=0, alpha=0)
ax.fill_between([0, 72], 25, 35, color=AMBER, alpha=0.12)
ax.plot([0, 72], [4, 4], "--", color=GREEN, lw=1.8, label="出口ストリーム品・ドラム/SP（規格内）")
ax.axhline(10, color=GRAY, ls=":", lw=1.4)
ax.text(2, 10.6, "検査規格 APHA10", fontproperties=F(9), color=GRAY)
ax.text(60, 32.5, "25〜35で頭打ち", fontproperties=F(9), color=AMBER)
ax.set_xlim(0, 72); ax.set_ylim(0, 38)
ax.set_xlabel("保管日数（日）", fontproperties=F(10)); ax.set_ylabel("APHA", fontproperties=F(10))
ax.set_title("製品DEGの色相（APHA）経時：タンクで上昇、出口/大気接触は不変", fontproperties=F(11, True), color=DBLUE)
ax.legend(prop=F(9), loc="upper left", framealpha=0.9)
ax.grid(alpha=0.25)
for lab in ax.get_xticklabels() + ax.get_yticklabels(): lab.set_fontproperties(F(9))
save(fig, "apha_trend.png")

# ── 2) UVスペクトル模式（300/450 vs 付着物380） ──
fig, ax = plt.subplots(figsize=(6.4, 3.2))
x = np.linspace(250, 700, 400)
def peak(c, w, h): return h * np.exp(-((x - c) ** 2) / (2 * w ** 2))
tankc = peak(300, 38, 0.9) + peak(450, 45, 0.55) + peak(600, 40, 0.12) + 0.05
deposit = peak(380, 42, 0.6) + 0.04
ax.plot(x, tankc, color=BLUE, lw=2.3, label="着色タンク品（実機）")
ax.plot(x, deposit, color=AMBER, lw=2.0, ls="--", label="C-1503付着物＋DEG（別物）")
ax.axvline(450, color=GREEN, ls=":", lw=1.3); ax.text(454, 0.82, "450nm＝黄", fontproperties=F(9), color=GREEN)
ax.axvline(380, color=AMBER, ls=":", lw=1.0); ax.text(330, 0.66, "380nm", fontproperties=F(9), color=AMBER)
ax.set_xlim(250, 700); ax.set_ylim(0, 1.05)
ax.set_xlabel("波長 (nm)", fontproperties=F(10)); ax.set_ylabel("吸光度（模式）", fontproperties=F(10))
ax.set_title("実機の着色は300/450nm。付着物は380nmのみ＝別現象", fontproperties=F(11, True), color=DBLUE)
ax.legend(prop=F(9), loc="upper right", framealpha=0.9)
ax.grid(alpha=0.2)
for lab in ax.get_xticklabels() + ax.get_yticklabels(): lab.set_fontproperties(F(9))
save(fig, "uv_spectrum.png")

# ── 3) UV450 早期判定 ──
fig, ax = plt.subplots(figsize=(6.4, 3.0))
dd = np.arange(0, 8)
bad = 0.002 + 0.004 * dd
good = 0.002 + 0.0 * dd
ax.plot(dd, bad, "-o", color=BLUE, lw=2.2, ms=5, label="悪化品（約0.004/日 上昇）")
ax.plot(dd, good, "-s", color=GREEN, lw=2.0, ms=4, label="正常品（0.002で不動）")
ax.fill_between(dd, good - 0.002, good + 0.002, color=GRAY, alpha=0.15)
ax.text(0.3, 0.0042, "測定精度 ±0.002", fontproperties=F(8.5), color=GRAY)
ax.set_xlim(0, 7.3); ax.set_ylim(0, 0.032)
ax.set_xlabel("経過日数（日）", fontproperties=F(10)); ax.set_ylabel("UV 450nm 吸光度", fontproperties=F(10))
ax.set_title("早期判定：UV450nmの日次の傾きで1日後から異常を判別", fontproperties=F(11, True), color=DBLUE)
ax.legend(prop=F(9), loc="upper left", framealpha=0.9)
ax.grid(alpha=0.25)
for lab in ax.get_xticklabels() + ax.get_yticklabels(): lab.set_fontproperties(F(9))
save(fig, "uv450_trend.png")
