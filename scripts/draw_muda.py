# -*- coding: utf-8 -*-
"""業務効率化／無駄の見える化 ── 1枚サマリPNG（無駄の型リスト版）。

配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）で content 系と統一。
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

f_title = F(16); f_sub = F(9); f_band = F(11); f_lbl = F(9); f_body = F(9)

PAD = 16
W = 1040
LH = 15
LBL_W = 150          # 型名列
STATE_W = 330        # 状態列
EX_W = W - PAD * 2 - LBL_W - STATE_W  # 例＋対策列

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


# ---- 行データ（型ごと）----
rows = []
for label, cat, state, examples, counter in C.MUDA_TYPES:
    ex = "例：" + " ／ ".join(examples)
    co = "→対策：" + counter
    rh = max(block_h(label + "\n(" + cat + ")", f_lbl, LBL_W - 12),
             block_h(state, f_body, STATE_W - 12),
             block_h(ex + "\n" + co, f_body, EX_W - 12), 34)
    rows.append((label, cat, state, ex, co, rh))

intro_text = "・" + "\n・".join(C.INTRO)
intro_h = block_h(intro_text, f_body, W - PAD * 2 - 12)
common_h = block_h(C.COMMON, f_body, W - PAD * 2 - 12)

title_h, sub_h, gap, band_h = 26, 26, 6, 24
H = (PAD + title_h + sub_h + 6 + band_h + intro_h + gap
     + band_h + 24 + sum(r[5] for r in rows)  # 型ヘッダ + 行
     + gap + band_h + common_h + PAD)

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


y = PAD
d.text((PAD * S, y * S), C.TITLE, font=f_title, fill=INK)
y += title_h
text_block(PAD, y, W - PAD * 2, sub_h, C.SUBTITLE, f_sub, SUB_INK)
y += sub_h
d.line([(PAD * S, y * S), ((W - PAD) * S, y * S)], fill=RULE, width=2)
y += 6

# 冒頭の考え
rect(PAD, y, W - PAD, y + band_h, HEADER)
text_block(PAD, y, W - PAD * 2, band_h, "無駄とは何か（私の考え）", f_band, INK, valign="center", bold=True)
y += band_h
rect(PAD, y, W - PAD, y + intro_h, WHITE)
text_block(PAD, y, W - PAD * 2 - 6, intro_h, intro_text, f_body, INK)
y += intro_h + gap

# 型一覧バンド
rect(PAD, y, W - PAD, y + band_h, HEADER)
text_block(PAD, y, W - PAD * 2, band_h, "自分の業務に見る“無駄の型”（過去〜直近・複数PJ／1個に絞らない）", f_band, INK, valign="center", bold=True)
y += band_h
# 列見出し
x = PAD
for w, t in [(LBL_W, "型／カテゴリ"), (STATE_W, "どういう状態が無駄か"), (EX_W, "実体験の例 ＋ 対策の方向")]:
    rect(x, y, x + w, y + 24, LABEL)
    text_block(x, y, w, 24, t, f_lbl, INK, align="center", valign="center", bold=True)
    x += w
y += 24
# 行
for label, cat, state, ex, co, rh in rows:
    x = PAD
    rect(x, y, x + LBL_W, y + rh, LABEL)
    text_block(x, y, LBL_W, rh, label + "\n(" + cat + ")", f_lbl, INK, valign="center", bold=True)
    x += LBL_W
    rect(x, y, x + STATE_W, y + rh, WHITE)
    text_block(x, y, STATE_W, rh, state, f_body, INK)
    x += STATE_W
    rect(x, y, x + EX_W, y + rh, WHITE)
    text_block(x, y, EX_W, rh, ex + "\n" + co, f_body, SUB_INK)
    y += rh
y += gap

# 共通点
rect(PAD, y, W - PAD, y + band_h, HEADER)
text_block(PAD, y, W - PAD * 2, band_h, "強いて共通点を言えば（※1個直せば終わり、ではない）", f_band, INK, valign="center", bold=True)
y += band_h
rect(PAD, y, W - PAD, y + common_h, WHITE)
text_block(PAD, y, W - PAD * 2 - 6, common_h, C.COMMON, f_body, INK)
y += common_h

path = os.path.join(OUT, "無駄の見える化_1枚サマリ.png")
img.save(path)
print("saved:", path, img.size)
