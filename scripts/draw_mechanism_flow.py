# -*- coding: utf-8 -*-
"""メカニズムの工程フロー図（横帯）。スライド4用。配布クリーン。"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrow
import matplotlib.font_manager as fm

fp = fm.FontProperties(fname="/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf")
NAVY = "#1F3864"; BLUE = "#2E5BA8"; LIGHT = "#EAEFF7"; GREEN = "#2E7D32"; RED = "#C0392B"; GRAY = "#555555"

OUT = os.path.join(os.path.dirname(__file__), "..", "outputs", "mechanism_flow.png")

steps = [
    ("EO反応器", "①アルデヒド増加\n（触媒選択性低下）", NAVY),
    ("脱水塔\n（低水分）", "低水分で\n縮合が進みやすい", BLUE),
    ("MEG塔", "縮合体（高沸点）\nとしてDEG留分へ", BLUE),
    ("DEG塔", "②アルドール縮合\n前駆体（MW150・330nm）", RED),
    ("製品タンク\n（N2・長期）", "③前駆体が熟成\n共役伸長→450nm", GREEN),
]

fig, ax = plt.subplots(figsize=(13.0, 3.15), dpi=200)
ax.set_xlim(0, 100); ax.set_ylim(0, 32); ax.axis("off")

n = len(steps); bw = 16.5; gap = (100 - n*bw) / (n+1)
x = gap; yb = 16; bh = 11
centers = []
for (name, mech, col) in steps:
    box = FancyBboxPatch((x, yb), bw, bh, boxstyle="round,pad=0.3,rounding_size=1.2",
                         linewidth=2, edgecolor=col, facecolor=LIGHT)
    ax.add_patch(box)
    ax.text(x+bw/2, yb+bh-2.4, name, ha="center", va="top", fontproperties=fp,
            fontsize=12.5, fontweight="bold", color=col)
    ax.text(x+bw/2, yb+bh-6.2, mech, ha="center", va="top", fontproperties=fp,
            fontsize=9.3, color="#222222")
    centers.append(x+bw/2)
    x += bw + gap

# arrows between
for i in range(n-1):
    xs = centers[i] + bw/2; xe = centers[i+1] - bw/2
    ax.add_patch(FancyArrow(xs+0.3, yb+bh/2, (xe-xs)-0.9, 0, width=0.5, head_width=2.2,
                            head_length=1.6, length_includes_head=True, color=BLUE))

# bottom band: oxygen effect
ax.add_patch(FancyBboxPatch((3, 8.6), 94, 5.4, boxstyle="round,pad=0.2,rounding_size=0.8",
                            linewidth=1.2, edgecolor=GRAY, facecolor="#F5F5F5", alpha=0.6))
ax.text(50, 11.5, "④ 酸素があると前駆体が有機酸へ酸化され縮合が進みにくい：窒素タンクで進行／大気接触（ドラム・SP）で抑制",
        ha="center", va="center", fontproperties=fp, fontsize=10.2, color=GRAY)

# top caption
ax.text(50, 30, "触媒劣化で増えたアルデヒド → アルドール縮合で高沸点の前駆体になりDEG留分へ → タンクで熟成して450nm（約10ppb・推定）",
        ha="center", va="center", fontproperties=fp, fontsize=11, fontweight="bold", color=NAVY)

plt.tight_layout(pad=0.4)
os.makedirs(os.path.dirname(OUT), exist_ok=True)
plt.savefig(OUT, bbox_inches="tight", facecolor="white")
print("saved:", os.path.normpath(OUT))
