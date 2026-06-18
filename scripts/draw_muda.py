# -*- coding: utf-8 -*-
"""業務効率化／無駄の見える化 ── 1枚サマリPNG（段階構成版）。

木村さんのお題 → 定義を開く → 自分ごと（型・基盤がない業務） → 対策の方向。
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）。
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
INK = (0, 0, 0); SUB_INK = (64, 64, 64); LINE = (128, 128, 128); RULE = (89, 89, 89)

f_title = F(16); f_sub = F(9); f_band = F(11); f_lbl = F(9); f_body = F(9)

PAD = 16
W = 1040
LH = 15
AREA_W = 150
STATE_W = 430
CO_W = W - PAD * 2 - AREA_W - STATE_W

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


odai = "・" + "\n・".join(C.ODAI)
defo = "・" + "\n・".join(C.DEF_OPEN)
tais = "・" + "\n・".join(C.TAISAKU)
odai_h = block_h(odai, f_body, W - PAD * 2 - 12)
defo_h = block_h(defo, f_body, W - PAD * 2 - 12)
tais_h = block_h(tais, f_body, W - PAD * 2 - 12)

rows = []
for area, state, counter in C.JIBUN:
    rh = max(block_h(area, f_lbl, AREA_W - 12),
             block_h(state, f_body, STATE_W - 12),
             block_h(counter, f_body, CO_W - 12), 30)
    rows.append((area, state, counter, rh))

title_h, sub_h, gap, band_h, hdr_h = 26, 26, 6, 24, 22
H = (PAD + title_h + sub_h + 6
     + band_h + odai_h + gap
     + band_h + defo_h + gap
     + band_h + hdr_h + sum(r[3] for r in rows) + gap
     + band_h + tais_h + PAD)

img = Image.new("RGB", (W * S, H * S), WHITE)
d = ImageDraw.Draw(img)


def rect(x0, y0, x1, y1, fillc, outline=LINE):
    d.rectangle([x0 * S, y0 * S, x1 * S, y1 * S], fill=fillc, outline=outline, width=1)


def text_block(x, y0, w, h, text, font, color, align="left", valign="top", bold=False):
    lines = wrap(d, text, font, w)
    th = len(lines) * LH
    ty = y0 + (h - th) / 2 if valign == "center" else y0 + 5
    for ln in lines:
        tw = d.textlength(ln, font=font)
        tx = x + (w - tw / S) / 2 if align == "center" else x + 6
        d.text((tx * S, ty * S), ln, font=font, fill=color)
        if bold:
            d.text((tx * S + 1, ty * S), ln, font=font, fill=color)
        ty += LH


def draw_band(y, text):
    rect(PAD, y, W - PAD, y + band_h, HEADER)
    text_block(PAD, y, W - PAD * 2, band_h, text, f_band, INK, valign="center", bold=True)
    return y + band_h


def draw_bullets(y, text, h):
    rect(PAD, y, W - PAD, y + h, WHITE)
    text_block(PAD, y, W - PAD * 2 - 6, h, text, f_body, INK)
    return y + h


y = PAD
d.text((PAD * S, y * S), C.TITLE, font=f_title, fill=INK)
y += title_h
text_block(PAD, y, W - PAD * 2, sub_h, C.SUBTITLE, f_sub, SUB_INK)
y += sub_h
d.line([(PAD * S, y * S), ((W - PAD) * S, y * S)], fill=RULE, width=2)
y += 6

y = draw_band(y, "木村さんのお題（6/17 朝会）")
y = draw_bullets(y, odai, odai_h) + gap
y = draw_band(y, "定義をもう少し開くと")
y = draw_bullets(y, defo, defo_h) + gap

y = draw_band(y, "自分ごとに落とすと：再現性がない＝「型・基盤」がなく毎回ゼロから")
x = PAD
for w, t in [(AREA_W, "領域"), (STATE_W, "どういう状態が無駄か（型・基盤がない）"), (CO_W, "対策の方向")]:
    rect(x, y, x + w, y + hdr_h, LABEL)
    text_block(x, y, w, hdr_h, t, f_lbl, INK, align="center", valign="center", bold=True)
    x += w
y += hdr_h
for area, state, counter, rh in rows:
    x = PAD
    rect(x, y, x + AREA_W, y + rh, LABEL)
    text_block(x, y, AREA_W, rh, area, f_lbl, INK, valign="center", bold=True)
    x += AREA_W
    rect(x, y, x + STATE_W, y + rh, WHITE)
    text_block(x, y, STATE_W, rh, state, f_body, SUB_INK)
    x += STATE_W
    rect(x, y, x + CO_W, y + rh, WHITE)
    text_block(x, y, CO_W, rh, counter, f_body, INK)
    y += rh
y += gap

y = draw_band(y, "対策の方向（再現性を上げる）")
y = draw_bullets(y, tais, tais_h)

path = os.path.join(OUT, "無駄の見える化_1枚サマリ.png")
img.save(path)
print("saved:", path, img.size)
