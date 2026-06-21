# -*- coding: utf-8 -*-
"""化学式の反応スキーム図（RDKit）。アルドール縮合で共役が伸び黄色になる筋を構造式で。淡色・見やすく。"""
import os
from rdkit import Chem
from rdkit.Chem.Draw import rdMolDraw2D
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import matplotlib.font_manager as fm
from PIL import Image
import io

OUTDIR = os.path.join(os.path.dirname(__file__), "..", "outputs")
fp = fm.FontProperties(fname="/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf")
BLUE = "#005BAB"; DBLUE = "#003F7E"; GRAY = "#6F6F6F"; GREEN = "#2E7D32"; AMBER = "#E0A000"

def mol_png(smiles, w=320, h=170, highlight_color=None):
    mol = Chem.MolFromSmiles(smiles)
    d = rdMolDraw2D.MolDraw2DCairo(w, h)
    opts = d.drawOptions()
    opts.bondLineWidth = 2
    opts.clearBackground = False
    rdMolDraw2D.PrepareAndDrawMolecule(d, mol)
    d.FinishDrawing()
    return Image.open(io.BytesIO(d.GetDrawingText()))

def place(ax, img, x, y, zoom=0.5):
    ab = AnnotationBbox(OffsetImage(img, zoom=zoom), (x, y), frameon=False)
    ax.add_artist(ab)

# ── 反応スキーム図 ──
fig, ax = plt.subplots(figsize=(13.0, 4.2), dpi=200)
ax.set_xlim(0, 100); ax.set_ylim(0, 32); ax.axis("off")

# molecules
acet = mol_png("CC=O", 240, 150)                         # アセトアルデヒド
croton = mol_png("C/C=C/C=O", 300, 150)                  # クロトンアルデヒド (共役2)
short = mol_png("C/C=C/C=C/C=C/C=O", 420, 150)           # 短い前駆体 (共役4) 330nm
long = mol_png("C/C=C/C=C/C=C/C=C/C=C/C=C/C=C/C=O", 640, 150)  # 長いポリエナール 450nm

ys = 19
place(ax, acet, 9, ys, 0.42)
place(ax, croton, 32, ys, 0.42)
place(ax, short, 58, ys, 0.40)
place(ax, long, 86, ys, 0.34)

# labels under
def lab(x, name, sub, col="#222222"):
    ax.text(x, 9.5, name, ha="center", va="top", fontproperties=fp, fontsize=10.5, fontweight="bold", color=col)
    ax.text(x, 6.2, sub, ha="center", va="top", fontproperties=fp, fontsize=9, color=GRAY)

lab(9, "アセトアルデヒド", "原料（種）")
lab(32, "クロトンアルデヒド", "共役2")
lab(58, "短い前駆体", "共役4・UV330nm（淡）")
lab(86, "ポリエナール", "共役を伸ばす・450nm（黄）", col=AMBER)

# arrows with conditions
def arrow(x0, x1, top, bot):
    ax.annotate("", xy=(x1, ys), xytext=(x0, ys),
                arrowprops=dict(arrowstyle="-|>", color=BLUE, lw=2.2))
    xm = (x0 + x1) / 2
    ax.text(xm, ys + 4.6, top, ha="center", va="bottom", fontproperties=fp, fontsize=8.6, color=DBLUE)
    ax.text(xm, ys - 5.2, bot, ha="center", va="top", fontproperties=fp, fontsize=8.6, color=GRAY)

arrow(15, 26, "アルドール縮合", "＋脱水（−H2O）")
arrow(40, 49, "逐次縮合", "共役伸長")
arrow(67, 76, "タンクで熟成", "前駆体どうしが縮合")

# top caption
ax.text(50, 30.5, "アルドール縮合と脱水で共役二重結合が伸びるほど吸収が長波長化 → 十分伸びると450nm（黄）",
        ha="center", va="center", fontproperties=fp, fontsize=11, fontweight="bold", color=BLUE)
# general formula
ax.text(50, 1.5, "一般式：  CH3-(CH=CH)n-CHO    （n が大きいほど黄色＝共役系のため微量で発色）",
        ha="center", va="center", fontproperties=fp, fontsize=10, color="#222222")

plt.tight_layout(pad=0.3)
os.makedirs(OUTDIR, exist_ok=True)
p = os.path.join(OUTDIR, "aldol_scheme.png")
plt.savefig(p, bbox_inches="tight", facecolor="white"); print("saved:", os.path.normpath(p))

# ── 長鎖ポリエナール単体（着色物質の代表構造） ──
fig2, ax2 = plt.subplots(figsize=(5.0, 2.2), dpi=200)
ax2.set_xlim(0, 100); ax2.set_ylim(0, 44); ax2.axis("off")
img = mol_png("C/C=C/C=C/C=C/C=C/C=C/C=C/C=C/C=C/C=C/C=O", 760, 200)
place(ax2, img, 50, 26, 0.34)
ax2.text(50, 6, "ポリエナール（共役を伸ばした例）", ha="center", va="center", fontproperties=fp,
         fontsize=10, fontweight="bold", color=AMBER)
plt.tight_layout(pad=0.2)
p2 = os.path.join(OUTDIR, "polyenal_long.png")
plt.savefig(p2, bbox_inches="tight", facecolor="white"); print("saved:", os.path.normpath(p2))
