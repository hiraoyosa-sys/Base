# -*- coding: utf-8 -*-
"""DEG着色シナリオ 横フロー整理（詳細版）PNG生成。

横フロー4段＋リン酸の効き場所＋一本化した筋＋機構の確定事項＋鉄の2つの顔。
チャット内に直表示できる読みやすい版（PIL, scale=2）。
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

TITLE   = (31, 56, 100)
HEAD    = (46, 84, 150)
HEAD_HL = (158, 58, 38)
KEPT    = (221, 235, 247)
REASON  = (226, 239, 218)
DROP    = (242, 242, 242)
DROPWHY = (217, 217, 217)
SECT    = (68, 84, 106)
IRON_OK = (226, 239, 218)
IRON_NG = (252, 228, 214)
WHITE   = (255, 255, 255)
INK     = (34, 34, 34)
GREY    = (89, 89, 89)
LINE    = (191, 191, 191)

f_title = F(20); f_date = F(11); f_head = F(12); f_lbl = F(11)
f_body = F(10); f_sect = F(13); f_suji = F(11); f_kv = F(10)

PAD = 14
LBL_W = 130
STEP_W = 248
NSTEP = len(C.STEPS)
W = PAD * 2 + LBL_W + STEP_W * NSTEP
LH = 15

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


def block_h(text, font, maxw, padv=10):
    return len(wrap(_d, text, font, maxw)) * LH + padv * 2


# ---- 高さ計算 ----
HL = [hl for _, hl in C.STEPS]
STEPS_TXT = [s for s, _ in C.STEPS]

tier_h = []
for t in range(4):
    h = max(block_h(C.DATA[i][t], f_body, STEP_W - 16) for i in range(NSTEP))
    h = max(h, block_h(C.ROW_LABELS[t], f_lbl, LBL_W - 12))
    tier_h.append(max(h, 40))

head_h = max(max(block_h(s, f_head, STEP_W - 12, padv=6) for s in STEPS_TXT), 46)

KV_LBL_W = 132
kv_inner = W - PAD * 2 - KV_LBL_W - 16


def kv_h(label, text):
    return max(block_h(text, f_kv, kv_inner), block_h(label, f_lbl, KV_LBL_W - 12), 30)


phos_h = [kv_h(k, v) for k, v in C.PHOS]
suji_h = block_h(C.SUJI, f_suji, W - PAD * 2 - 16, padv=10)
mech_h = [kv_h(k, v) for k, v in C.MECH]
iron_h = [kv_h(k, v) for k, v in C.IRON]

SEC = 26   # セクション帯高
GAP = 14

title_h, date_h = 34, 18
H = (PAD + title_h + date_h + head_h + sum(tier_h)
     + GAP + SEC + sum(phos_h)
     + GAP + SEC + suji_h
     + GAP + SEC + sum(mech_h)
     + GAP + SEC + sum(iron_h)
     + PAD)

img = Image.new("RGB", (W * S, H * S), WHITE)
d = ImageDraw.Draw(img)


def rect(x0, y0, x1, y1, fillc, outline=LINE):
    d.rectangle([x0 * S, y0 * S, x1 * S, y1 * S], fill=fillc, outline=outline, width=1)


def text_block(x, y0, w, h, text, font, color, align="left", valign="top", bold=False):
    lines = wrap(d, text, font, w)
    th = len(lines) * LH
    ty = y0 + (h - th) / 2 if valign == "center" else y0 + 8
    for ln in lines:
        tw = d.textlength(ln, font=font)
        tx = x + (w - tw / S) / 2 if align == "center" else x + 6
        d.text((tx * S, ty * S), ln, font=font, fill=color)
        if bold:
            d.text((tx * S + 1, ty * S), ln, font=font, fill=color)
        ty += LH


def section(y, label):
    rect(PAD, y, W - PAD, y + SEC, SECT, outline=SECT)
    d.text(((PAD + 8) * S, (y + 5) * S), label, font=f_sect, fill=WHITE)
    return y + SEC


def kv_rows(y, items, heights, lblfill=DROP, valfill=WHITE, lblfills=None, valfills=None):
    for i, ((k, v), hh) in enumerate(zip(items, heights)):
        lf = lblfills[i] if lblfills else lblfill
        vf = valfills[i] if valfills else valfill
        rect(PAD, y, PAD + KV_LBL_W, y + hh, lf)
        text_block(PAD, y, KV_LBL_W, hh, k, f_lbl, (51, 51, 51), align="center", valign="center", bold=True)
        rect(PAD + KV_LBL_W, y, W - PAD, y + hh, vf)
        text_block(PAD + KV_LBL_W, y, kv_inner, hh, v, f_kv, INK, valign="center")
        y += hh
    return y


# ---- 描画 ----
y = PAD
rect(PAD, y, W - PAD, y + title_h, TITLE, outline=TITLE)
d.text(((PAD + 8) * S, (y + 6) * S), C.TITLE, font=f_title, fill=WHITE)
y += title_h
tw = d.textlength(C.DATE, font=f_date)
d.text(((W - PAD - 6) * S - tw, (y + 3) * S), C.DATE, font=f_date, fill=GREY)
y += date_h

# 工程ヘッダ
x = PAD
rect(x, y, x + LBL_W, y + head_h, SECT, outline=SECT)
text_block(x, y, LBL_W, head_h, "工程フロー →", f_head, WHITE, align="center", valign="center")
x += LBL_W
for i, s in enumerate(STEPS_TXT):
    col = HEAD_HL if HL[i] else HEAD
    rect(x, y, x + STEP_W, y + head_h, col, outline=col)
    text_block(x, y, STEP_W, head_h, s, f_head, WHITE, align="center", valign="center", bold=True)
    if i < NSTEP - 1:
        ax = x + STEP_W; ay = y + head_h / 2
        d.polygon([((ax - 5) * S, (ay - 5) * S), ((ax + 4) * S, ay * S), ((ax - 5) * S, (ay + 5) * S)], fill=WHITE)
    x += STEP_W
y += head_h

# 4段
ROW_FILLS = [KEPT, REASON, DROP, DROPWHY]
for t in range(4):
    h = tier_h[t]; x = PAD
    rect(x, y, x + LBL_W, y + h, ROW_FILLS[t])
    text_block(x, y, LBL_W, h, C.ROW_LABELS[t], f_lbl, (51, 51, 51), valign="center", bold=True)
    x += LBL_W
    for i in range(NSTEP):
        rect(x, y, x + STEP_W, y + h, ROW_FILLS[t])
        text_block(x, y, STEP_W - 4, h, C.DATA[i][t], f_body, INK, valign="top")
        x += STEP_W
    y += h

# リン酸
y += GAP
y = section(y, "リン酸の効き場所（絞り込み済み・確定未）")
y = kv_rows(y, C.PHOS, phos_h, lblfill=DROP)

# 一本化
y += GAP
y = section(y, "一本化した筋")
rect(PAD, y, W - PAD, y + suji_h, REASON)
text_block(PAD, y, W - PAD * 2, suji_h, C.SUJI, f_suji, INK, valign="center")
y += suji_h

# 機構の確定事項
y += GAP
y = section(y, "機構の確定事項（理由欄の裏付け）")
y = kv_rows(y, C.MECH, mech_h, lblfill=KEPT)

# 鉄の2つの顔
y += GAP
y = section(y, "鉄の役割＝2つの顔（今回の深掘り）")
iron_lf = [IRON_OK, IRON_NG] + [DROP] * (len(C.IRON) - 2)
iron_vf = [IRON_OK, IRON_NG] + [WHITE] * (len(C.IRON) - 2)
y = kv_rows(y, C.IRON, iron_h, lblfills=iron_lf, valfills=iron_vf)

path = os.path.join(OUT, "DEG着色_フロー整理.png")
img.save(path)
print("saved:", path, img.size)
