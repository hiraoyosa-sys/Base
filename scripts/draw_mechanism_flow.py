# -*- coding: utf-8 -*-
"""メカニズムの工程フロー図（横帯）。スライド4用。配布クリーン。"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrow
import matplotlib.font_manager as fm

fp = fm.FontProperties(fname="/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf")
# 薄いMC配色（濃色は使わない。テキスト/枠線にMCブルー、塗りは淡色）
NAVY = "#003F7E"; BLUE = "#005BAB"; LIGHT = "#EBEFF2"; GREEN = "#005BAB"; RED = "#005BAB"; GRAY = "#666666"
TANK = "#FFF6E0"  # 着色の場（タンク）を淡いアンバーで強調

OUT = os.path.join(os.path.dirname(__file__), "..", "outputs", "mechanism_flow.png")

steps = [
    ("EO反応系", "①アルデヒド\n増加", NAVY),
    ("EG反応系", "グリコールと\n共存", BLUE),
    ("EG濃縮系", "大半パージ\n一部下流へ", BLUE),
    ("EG脱水系", "低水分で\n縮合が進む", BLUE),
    ("MEG精製系", "通過点\n（金属なし）", BLUE),
    ("DEG精製系", "②前駆体生成\n330nm・MW150", RED),
    ("製品タンク\n(N2・長期)", "③熟成→共役\n伸長→450nm", GREEN),
]

fig, ax = plt.subplots(figsize=(13.0, 3.15), dpi=200)
ax.set_xlim(0, 100); ax.set_ylim(0, 32); ax.axis("off")

n = len(steps); gap = 1.6; bw = (100 - (n+1)*gap) / n
x = gap; yb = 16; bh = 11
centers = []
for (name, mech, col) in steps:
    fc = TANK if "タンク" in name else LIGHT
    box = FancyBboxPatch((x, yb), bw, bh, boxstyle="round,pad=0.3,rounding_size=1.2",
                         linewidth=2, edgecolor=col, facecolor=fc)
    ax.add_patch(box)
    ax.text(x+bw/2, yb+bh-2.0, name, ha="center", va="top", fontproperties=fp,
            fontsize=10.0, fontweight="bold", color=col)
    ax.text(x+bw/2, yb+bh-5.8, mech, ha="center", va="top", fontproperties=fp,
            fontsize=8.2, color="#222222")
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
