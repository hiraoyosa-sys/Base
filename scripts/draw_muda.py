# -*- coding: utf-8 -*-
"""業務効率化／無駄の見える化 ── 1枚サマリPNG生成。

中心命題＋再現性4要件＋ロジックラインを1枚に圧縮。
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）で content.py 系と統一。
"""
import os
from PIL import Image, ImageDraw, ImageFont
import content_muda as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)
FONT_PATH = "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"

S = 2
def F(sz):
    return ImageFont.truetype(FONT_PATH, sz * S)

WHITE = (255, 255, 255); HEADER = (230, 230, 230); LABEL = (242, 242, 242)
SUB = (247, 247, 247); INK = (0, 0, 0); SUB_INK = (64, 64, 64)
LINE = (128, 128, 128); RULE = (89, 89, 89)

f_title = F(17); f_sub = F(9); f_band = F(11); f_lbl = F(9); f_body = F(9)

PAD = 16
W = 980
LH = 15

_tmp = Image.new("RGB", (10, 10)); _d = ImageDraw.Draw(_tmp)


def wrap(draw, text, font, maxw):
    out = []
    for para in str(text).split("\n"):
        if para == "":
            out.append(""); continue
        line = ""
        for ch in para:
            if draw.textlength(line + ch, font=font) <= maxw * S:
                line += ch
            else:
                out.append(line); line = ch
        out.append(line)
    return out


def block_h(text, font, maxw, padv=7):
    return len(wrap(_d, text, font, maxw)) * LH + padv * 2


# 高さを先に積算
ROWS = []  # (kind, payload)
ROWS.append(("band", "中心命題 ── 無駄とは“労力を二度払う状態”"))
ROWS.append(("para", C.THESIS))
ROWS.append(("band", "スコープ（扱う／扱わない）"))
ROWS.append(("para_sub", C.SCOPE))
ROWS.append(("band", "無駄になりやすい4つの軸（どれかが欠けると作り直しになる）"))
for req, one, lack, mine in C.REQS:
    ROWS.append(("req", (req, one, lack)))
ROWS.append(("band", "上位目的への接続（残業削減は結果指標）"))
ROWS.append(("logic", "　→　".join(C.LOGIC)))
ROWS.append(("band", "最初の1アクション"))
ROWS.append(("para", C.FIRST_ACTION))

inner = W - PAD * 2
title_h, sub_h, gap = 28, 26, 6
total = PAD + title_h + sub_h + 6
heights = []
for kind, payload in ROWS:
    if kind == "band":
        h = 24
    elif kind in ("para", "logic"):
        h = block_h(payload, f_body, inner - 12)
    elif kind == "para_sub":
        h = block_h(payload, f_body, inner - 12)
    elif kind == "req":
        req, one, lack = payload
        h = max(block_h(one, f_body, 300 - 12), block_h(lack, f_body, inner - 150 - 300 - 12), 30)
    heights.append(h)
    total += h + gap
H = total + PAD

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
d.text((PAD * S, y * S), C.TITLE, font=f_title, fill=INK)
y += title_h
text_block(PAD, y, inner, sub_h, C.SUBTITLE, f_sub, SUB_INK, valign="top")
y += sub_h
d.line([(PAD * S, y * S), ((W - PAD) * S, y * S)], fill=RULE, width=2)
y += 6

for (kind, payload), h in zip(ROWS, heights):
    if kind == "band":
        rect(PAD, y, W - PAD, y + h, HEADER)
        text_block(PAD, y, inner, h, payload, f_band, INK, valign="center", bold=True)
    elif kind in ("para",):
        rect(PAD, y, W - PAD, y + h, WHITE)
        text_block(PAD, y, inner - 6, h, payload, f_body, INK)
    elif kind == "para_sub":
        rect(PAD, y, W - PAD, y + h, SUB)
        text_block(PAD, y, inner - 6, h, payload, f_body, SUB_INK)
    elif kind == "logic":
        rect(PAD, y, W - PAD, y + h, WHITE)
        text_block(PAD, y, inner - 6, h, payload, f_body, INK)
    elif kind == "req":
        req, one, lack = payload
        rect(PAD, y, PAD + 150, y + h, LABEL)
        text_block(PAD, y, 150, h, req + "　" + one, f_lbl, INK, valign="center", bold=True)
        rect(PAD + 150, y, W - PAD, y + h, WHITE)
        text_block(PAD + 150, y, inner - 150 - 6, h, "欠落すると：" + lack, f_body, SUB_INK, valign="center")
    y += h + gap

path = os.path.join(OUT, "無駄の見える化_1枚サマリ.png")
img.save(path)
print("saved:", path, img.size)
