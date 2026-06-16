# -*- coding: utf-8 -*-
"""DEG着色シナリオ 横フロー整理（4段）PNG生成。

横＝プロセスフロー、縦＝4段。下部にリン酸の効き場所と一本化した筋。
チャット内に直表示できる読みやすい版（PIL, scale=2）。
"""
import os
from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)
FONT_PATH = "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"

S = 2  # scale
def F(sz):
    return ImageFont.truetype(FONT_PATH, sz * S)

# 色
TITLE   = (31, 56, 100)
HEAD    = (46, 84, 150)
HEAD_HL = (158, 58, 38)
KEPT    = (221, 235, 247)
REASON  = (226, 239, 218)
DROP    = (242, 242, 242)
DROPWHY = (217, 217, 217)
SECT    = (68, 84, 106)
WHITE   = (255, 255, 255)
INK     = (34, 34, 34)
GREY    = (89, 89, 89)
LINE    = (191, 191, 191)

STEPS = ["反応系\n(EO反応器)", "EG濃縮系\n(パージの場)", "脱水塔\n(乾いた条件)",
         "MEG塔\n(C-1501/1502)", "DEG塔(C-1503)\n★着色の本体", "製品タンク\n(出荷・船・ローリー)"]
HL = [False, False, False, False, True, False]

DATA = [
    ["ALD(A/F-ALD)が増加。水が多く遊離のまま。",
     "EO触媒劣化で原料ALD増は事実一致。従来も微量(1ppm以下)常在。",
     "（なし）", "—"],
    ["大半はパージで系外へ。一部が隠れ形(アセタール等)で逃げ切る。",
     "A-ALD等はここでパージが通常。原料ALD増で逃げ切る量が増え顕在化、と整合(要検証)。",
     "「全量パージされ下流に来ない」を否定。",
     "実際に下流でALDが見えている。"],
    ["アセタール化で重質化し下流まで運ばれる。",
     "水が少なくアセタール側に寄る。重質化すれば届く。",
     "× 軽質ALDのままDEGへ抜ける。",
     "軽質ALDは物性上そのままDEGに来ない(高橋)。運ぶには重質化が必要。"],
    ["熱・アルカリで隠れ形が分解しALDに戻る(C-1502付近で湧いて見える)。",
     "バランス上C-1502付近でA/F-ALDが生成するように見える(C-1501ボトムは過去同水準＝MEG塔単独生成でない)。隠れ形分解で説明可。",
     "× MEG塔そのものが着色の場。",
     "C-1501/1502開放(6/12)→サビ・汚れなし。触媒(活性サビ)の場でない。"],
    ["高pH＋高温＋活性サビが揃い、アルドール縮合→脱水で共役ポリエナール(300/450nm)。塩基触媒・酸素フリー。",
     "①酸素フリーで進む＝酸素で止まる事実(F2)と整合 ②C-1503トップに活性サビ(黒異物・XRF98.5%鉄)実在 ③高pH＋高温が揃う唯一の塔 ④A-ALD自己縮合は共役3-4まで実測(9連結は微量だが微量で色がつく) ⑤活性サビ＝ルイス酸点。高pH(塩基)と組んだ酸塩基二元触媒で脱水(発色)を加速。酸素・電子移動不要でF2と両立し、高pH＋活性サビが揃う唯一塔=C-1503を補強。",
     "× 鉄＋酸素のラジカル重合が主役。\n△ ルートB(F+A交差→アクロレイン)は残す(要検証)。",
     "ラジカル重合は酸素が開始剤→酸素で止まる事実と逆、鉄サビ単独で着色せず(RD)。\nルートBは450nm直接証拠なし・寄与度不明で要検証(否定はせず)。\n※落とすのは「ラジカル開始剤としての鉄」。ルイス酸触媒としての鉄は残す(残した理由⑤)。"],
    ["残った前駆体が保管中に成長→450nm。塩基触媒で酸素なしに進む(縮合/アクロレインのアニオン重合)。",
     "本管DEGは27日でAPHA5以下・無変質、タンク品で450成長＝成長はタンク(滞留)で起きる。N2・遮光でも進行＝酸素不要。",
     "× タンクでラジカル重合して着色。",
     "酸素雰囲気(船・ローリー)で悪化が止まる事実と逆(ラジカル重合は酸素が開始剤)。"],
]
ROW_LABELS = ["残った仮説\n（本流）", "残した理由\n（なぜ生き残るか）", "落とした仮説",
              "落とした理由\n（事実紐付け）"]
ROW_FILLS = [KEPT, REASON, DROP, DROPWHY]

# レイアウト（論理px、最後に*Sしてfontと整合）
PAD = 14
LBL_W = 120
STEP_W = 232
NSTEP = len(STEPS)
W = PAD * 2 + LBL_W + STEP_W * NSTEP

f_title = F(20)
f_date = F(11)
f_head = F(12)
f_lbl = F(11)
f_body = F(10)
f_sect = F(13)
f_suji = F(11)

def wrap(draw, text, font, maxw):
    """改行コード尊重＋幅で折返し。返り値: 行リスト。"""
    out = []
    for para in text.split("\n"):
        if para == "":
            out.append("")
            continue
        line = ""
        for ch in para:
            test = line + ch
            if draw.textlength(test, font=font) <= maxw * S:
                line = test
            else:
                out.append(line)
                line = ch
        out.append(line)
    return out

# 計測用ダミー
_tmp = Image.new("RGB", (10, 10))
_d = ImageDraw.Draw(_tmp)

LH = 15  # 行高(論理px)
def block_h(text, font, maxw, padv=10):
    n = len(wrap(_d, text, font, maxw))
    return n * LH + padv * 2

# 各段の高さ＝その段で最も高いセル
tier_h = []
for t in range(4):
    h = max(block_h(DATA[i][t], f_body, STEP_W - 16) for i in range(NSTEP))
    h = max(h, block_h(ROW_LABELS[t], f_lbl, LBL_W - 12))
    tier_h.append(max(h, 40))

# ヘッダ高
head_h = max(block_h(s, f_head, STEP_W - 12, padv=6) for s in STEPS)
head_h = max(head_h, 46)

# 下部セクション
phos = [
    ("効き方", "リン酸鉄被膜が鉄表面サイトを封鎖→鉄の触媒作用を停止。redoxサイクル封じに限定せず、ルイス酸触媒(ルートA加速)としての鉄でも成立しF2と両立。"),
    ("ケースA", "C-1503でP検出 → DEG塔で金属(活性サビ)不活性化＝現行説成立。本線。"),
    ("ケースB", "C-1503でP無し → 効いた場所は上流(MEG/C-1502)で前駆体生成抑制（高橋説）。着色現場=DEG塔とテンション。"),
    ("決め手①", "C-1503ミドル充填部の壁面P分析（足場後採取）。リン入れてMEG-UV効かず=上流説の材料(F8)も突き合わせ。"),
    ("決め手②", "活性黒サビ有/無で酸素フリー着色速度を比較。加速ならルイス酸触媒を直接証明（redox・酸素から分離）。"),
]
suji = ("原料ALD増 → 大半パージ・一部隠れ形で逃げる → 脱水塔以降で重質化して運搬 → "
        "MEG塔系で分解しALDに戻り「湧いて見える」 → C-1503(高pH・高温・活性サビ＝酸塩基二元触媒)でアルドール縮合→共役ポリエナール → "
        "タンクで酸素なしに成長 → 450nm着色。\n"
        "動かない軸＝DEGに活性アルデヒドを入れない（IERのガードがなく活性ALDが残ると着色物質に化ける）。")

PHOS_LBL_W = 80
phos_h = [max(block_h(v, f_body, W - PAD*2 - PHOS_LBL_W - 16), 32) for _, v in phos]
suji_h = block_h(suji, f_suji, W - PAD*2 - 16, padv=10)

# 総高
y = PAD
title_h = 34
date_h = 18
H = (PAD + title_h + date_h + head_h + sum(tier_h)
     + 16 + 26 + sum(phos_h) + 16 + 26 + suji_h + PAD)

img = Image.new("RGB", (W * S, H * S), WHITE)
d = ImageDraw.Draw(img)

def rect(x0, y0, x1, y1, fill, outline=LINE):
    d.rectangle([x0*S, y0*S, x1*S, y1*S], fill=fill, outline=outline, width=1)

def text_block(x, y0, w, h, text, font, color, align="left", valign="top", bold=False):
    lines = wrap(d, text, font, w)
    th = len(lines) * LH
    if valign == "center":
        ty = y0 + (h - th) / 2
    else:
        ty = y0 + 8
    for ln in lines:
        tw = d.textlength(ln, font=font)
        if align == "center":
            tx = x + (w - tw / S) / 2
        else:
            tx = x + 6
        d.text((tx*S, ty*S), ln, font=font, fill=color)
        if bold:
            d.text((tx*S+1, ty*S), ln, font=font, fill=color)
        ty += LH

# タイトル
rect(PAD, y, W - PAD, y + title_h, TITLE, outline=TITLE)
d.text(((PAD+8)*S, (y+6)*S), "DEG着色シナリオ　横フロー整理", font=f_title, fill=WHITE)
y += title_h
# 日付
tw = d.textlength("2026-06-15", font=f_date)
d.text(((W - PAD - 6)*S - tw, (y+3)*S), "2026-06-15", font=f_date, fill=GREY)
y += date_h

# ヘッダ行
x = PAD
rect(x, y, x + LBL_W, y + head_h, SECT, outline=SECT)
text_block(x, y, LBL_W, head_h, "工程フロー →", f_head, WHITE, align="center", valign="center")
x += LBL_W
for i, s in enumerate(STEPS):
    col = HEAD_HL if HL[i] else HEAD
    rect(x, y, x + STEP_W, y + head_h, col, outline=col)
    text_block(x, y, STEP_W, head_h, s, f_head, WHITE, align="center", valign="center", bold=True)
    # 矢印
    if i < NSTEP - 1:
        ax = x + STEP_W
        ay = y + head_h / 2
        d.polygon([((ax-5)*S, (ay-5)*S), ((ax+4)*S, ay*S), ((ax-5)*S, (ay+5)*S)], fill=WHITE)
    x += STEP_W
y += head_h

# 4段
for t in range(4):
    h = tier_h[t]
    x = PAD
    rect(x, y, x + LBL_W, y + h, ROW_FILLS[t])
    text_block(x, y, LBL_W, h, ROW_LABELS[t], f_lbl, (51,51,51), align="left", valign="center", bold=True)
    x += LBL_W
    for i in range(NSTEP):
        rect(x, y, x + STEP_W, y + h, ROW_FILLS[t])
        text_block(x, y, STEP_W - 4, h, DATA[i][t], f_body, INK, align="left", valign="top")
        x += STEP_W
    y += h

y += 16
# リン酸セクション帯
rect(PAD, y, W - PAD, y + 26, SECT, outline=SECT)
d.text(((PAD+8)*S, (y+5)*S), "リン酸の効き場所（絞り込み済み・確定未）", font=f_sect, fill=WHITE)
y += 26
for (k, v), hh in zip(phos, phos_h):
    rect(PAD, y, PAD + PHOS_LBL_W, y + hh, DROP)
    text_block(PAD, y, PHOS_LBL_W, hh, k, f_lbl, (51,51,51), align="center", valign="center", bold=True)
    rect(PAD + PHOS_LBL_W, y, W - PAD, y + hh, WHITE)
    text_block(PAD + PHOS_LBL_W, y, W - PAD*2 - PHOS_LBL_W, hh, v, f_body, INK, valign="center")
    y += hh

y += 16
# 一本化セクション
rect(PAD, y, W - PAD, y + 26, SECT, outline=SECT)
d.text(((PAD+8)*S, (y+5)*S), "一本化した筋", font=f_sect, fill=WHITE)
y += 26
rect(PAD, y, W - PAD, y + suji_h, REASON)
text_block(PAD, y, W - PAD*2, suji_h, suji, f_suji, INK, valign="center")
y += suji_h

path = os.path.join(OUT, "DEG着色_フロー整理.png")
img.save(path)
print("saved:", path, img.size)
