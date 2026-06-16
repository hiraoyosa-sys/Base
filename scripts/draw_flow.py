# -*- coding: utf-8 -*-
"""DEG着色シナリオ 横フロー整理（最新版）PNG生成。

横＝プロセス工程、縦＝トピック行のマトリクス。各セルは列内に割り付け。
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）。
"""
import os
from PIL import Image, ImageDraw, ImageFont
import content as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)
FONT_PATH = "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"

S = 2
def F(sz):
    return ImageFont.truetype(FONT_PATH, sz * S)

WHITE   = (255, 255, 255)
HEADER  = (230, 230, 230)
LABEL   = (242, 242, 242)
SUB     = (247, 247, 247)
INK     = (0, 0, 0)
SUB_INK = (64, 64, 64)
LINE    = (128, 128, 128)
RULE    = (89, 89, 89)

f_title = F(18); f_date = F(10); f_head = F(11); f_lbl = F(9)
f_body = F(9); f_suji = F(10)

PAD = 14
LBL_W = 118
STEP_W = 250
NSTEP = len(C.STEPS)
W = PAD * 2 + LBL_W + STEP_W * NSTEP
LH = 14

_tmp = Image.new("RGB", (10, 10)); _d = ImageDraw.Draw(_tmp)


def wrap(draw, text, font, maxw):
    out = []
    for para in text.split("\n"):
        if para == "":
            out.append("")
            continue
        line = ""
        for ch in para:
            if draw.textlength(line + ch, font=font) <= maxw * S:
                line += ch
            else:
                out.append(line); line = ch
        out.append(line)
    return out


def block_h(text, font, maxw, padv=8):
    return len(wrap(_d, text, font, maxw)) * LH + padv * 2


SHADED = {"rust", "drug", "facts", "todo"}
STRONG = {"kept", "kept_reason"}

# 行高計算
row_h = []
for label, key in C.ROW_DEFS:
    h = max(block_h(C.COLS[i][key], f_body, STEP_W - 14) for i in range(NSTEP))
    h = max(h, block_h(label, f_lbl, LBL_W - 12))
    row_h.append(max(h, 30))

head_h = max(max(block_h(s, f_head, STEP_W - 12, padv=4) for s, _ in C.STEPS), 40)
suji_h = block_h(C.SUJI, f_suji, W - PAD * 2 - LBL_W - 16, padv=8)

title_h, date_h, gap = 30, 16, 10
H = PAD + title_h + date_h + head_h + sum(row_h) + gap + suji_h + PAD

img = Image.new("RGB", (W * S, H * S), WHITE)
d = ImageDraw.Draw(img)


def rect(x0, y0, x1, y1, fillc, outline=LINE):
    d.rectangle([x0 * S, y0 * S, x1 * S, y1 * S], fill=fillc, outline=outline, width=1)


def text_block(x, y0, w, h, text, font, color, align="left", valign="top", bold=False):
    lines = wrap(d, text, font, w)
    th = len(lines) * LH
    ty = y0 + (h - th) / 2 if valign == "center" else y0 + 6
    for ln in lines:
        tw = d.textlength(ln, font=font)
        tx = x + (w - tw / S) / 2 if align == "center" else x + 6
        d.text((tx * S, ty * S), ln, font=font, fill=color)
        if bold:
            d.text((tx * S + 1, ty * S), ln, font=font, fill=color)
        ty += LH


y = PAD
# タイトル（塗りなし・黒太字）
d.text((PAD * S, y * S), C.TITLE, font=f_title, fill=INK)
y += title_h
dt = f"（{C.DATE} 時点）"
tw = d.textlength(dt, font=f_date)
d.text(((W - PAD) * S - tw, (y + 2) * S), dt, font=f_date, fill=SUB_INK)
d.line([(PAD * S, (y + date_h) * S), ((W - PAD) * S, (y + date_h) * S)], fill=RULE, width=2)
y += date_h

# 工程ヘッダ
x = PAD
rect(x, y, x + LBL_W, y + head_h, HEADER)
text_block(x, y, LBL_W, head_h, "工程 →", f_head, INK, align="center", valign="center", bold=True)
x += LBL_W
for i, (name, hl) in enumerate(C.STEPS):
    rect(x, y, x + STEP_W, y + head_h, HEADER)
    text_block(x, y, STEP_W, head_h, name, f_head, INK, align="center", valign="center", bold=True)
    if i < NSTEP - 1:
        ax = x + STEP_W; ay = y + head_h / 2
        d.polygon([((ax - 5) * S, (ay - 5) * S), ((ax + 4) * S, ay * S), ((ax - 5) * S, (ay + 5) * S)], fill=INK)
    x += STEP_W
# ヘッダ下に強め罫線
d.line([(PAD * S, (y + head_h) * S), ((W - PAD) * S, (y + head_h) * S)], fill=RULE, width=2)
y += head_h

# トピック行
for ridx, (label, key) in enumerate(C.ROW_DEFS):
    h = row_h[ridx]
    shade = SUB if key in SHADED else WHITE
    x = PAD
    rect(x, y, x + LBL_W, y + h, LABEL)
    text_block(x, y, LBL_W, h, label, f_lbl, INK, valign="center", bold=True)
    x += LBL_W
    for i in range(NSTEP):
        rect(x, y, x + STEP_W, y + h, shade)
        color = INK if key in STRONG else SUB_INK
        text_block(x, y, STEP_W - 4, h, C.COLS[i][key], f_body, color, valign="top")
        x += STEP_W
    y += h

# 一本化した筋
y += gap
x = PAD
rect(x, y, x + LBL_W, y + suji_h, HEADER)
text_block(x, y, LBL_W, suji_h, "一本化した筋", f_lbl, INK, valign="center", bold=True)
rect(x + LBL_W, y, W - PAD, y + suji_h, WHITE)
text_block(x + LBL_W, y, W - PAD - LBL_W - PAD, suji_h, C.SUJI, f_suji, INK, valign="center")
y += suji_h

path = os.path.join(OUT, "DEG着色_フロー整理.png")
img.save(path)
print("saved:", path, img.size)
