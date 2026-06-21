# -*- coding: utf-8 -*-
"""対策会議PPTX：元テンプレート(MC)を土台に、薄色・図/化学式・工程別マトリクスで構築。配布クリーン。"""
import os
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import deg_master as M

HERE = os.path.dirname(__file__)
TMPL = "/root/.claude/uploads/0ac036e2-7cd8-5317-a99c-aa39021935be/7d006f34-__DEG________EO__.pptx"
OUTDIR = os.path.join(HERE, "..", "outputs")
OUT = os.path.join(OUTDIR, "製品DEG色相悪化_対策会議.pptx")
IN = 914400
EA = "游ゴシック"
# 薄いMC配色
BLUE = RGBColor(0x00, 0x5B, 0xAB)
DBLUE = RGBColor(0x00, 0x3F, 0x7E)
INK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x66, 0x66, 0x66)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
H_FILL = RGBColor(0xEB, 0xEF, 0xF2)   # 淡
PROC = RGBColor(0xD7, 0xE0, 0xE5)     # 淡青灰
C_OPE = RGBColor(0xE8, 0xF1, 0xE5)
C_RD = RGBColor(0xFB, 0xF1, 0xDC)
C_LIT = RGBColor(0xE9, 0xF0, 0xF7)
C_SUJI = RGBColor(0xDC, 0xE6, 0xEF)
AMBER = RGBColor(0xFF, 0xF6, 0xE0)
LINE = RGBColor(0xC9, 0xD2, 0xD8)

prs = Presentation(TMPL)
cp = prs.core_properties
cp.author = ""; cp.last_modified_by = ""; cp.title = "製品DEG色相悪化 対策会議"; cp.comments = ""

def layout(name):
    for l in prs.slide_layouts:
        if l.name == name: return l
    return prs.slide_layouts[0]

LY_COVER = layout("1_basic_cover")
LY_TEXT = layout("タイトルとコンテンツ") if any(l.name == "タイトルとコンテンツ" for l in prs.slide_layouts) else layout("1_basic_text")

def _ea(run, latin=EA):
    rPr = run._r.get_or_add_rPr()
    for tag, tf in (("a:latin", latin), ("a:ea", EA), ("a:cs", EA)):
        el = rPr.find(qn(tag))
        if el is None: el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", tf)

def newslide(ly):
    return prs.slides.add_slide(ly)

def drop_body_ph(s, keep_title=True):
    for ph in list(s.placeholders):
        idx = ph.placeholder_format.idx
        if idx != 0:  # keep title(0); drop others (content/date/number/footer left as-is by template)
            pass

def set_title(s, text):
    for ph in s.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.text = ""
            r = ph.text_frame.paragraphs[0].add_run(); r.text = text
            r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = DBLUE; _ea(r)
            return

def remove_content_ph(s):
    """本文プレースホルダ(idx!=0かつ番号/フッタ以外)を消し、自前で置く。"""
    for ph in list(s.placeholders):
        idx = ph.placeholder_format.idx
        if idx not in (0,) and ph.placeholder_format.type is not None:
            # フッタ/スライド番号/日付は残す（テンプレ装飾）
            from pptx.enum.shapes import PP_PLACEHOLDER
            if ph.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT,
                                              PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.CENTER_TITLE):
                ph._element.getparent().remove(ph._element)

def tbx(s, l, t, w, h):
    x = s.shapes.add_textbox(Emu(int(l*IN)), Emu(int(t*IN)), Emu(int(w*IN)), Emu(int(h*IN)))
    x.text_frame.word_wrap = True
    return x.text_frame

def P(tf, text, size, bold=False, first=False, color=INK, bullet=False, space=4, align=None):
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    if align is not None: p.alignment = align
    r = p.add_run(); r.text = ("・" + text) if bullet else text
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
    p.space_after = Pt(space); p.space_before = Pt(0)
    return p

def rrect(s, l, t, w, h, fill, line=None, lw=1.0):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(int(l*IN)), Emu(int(t*IN)), Emu(int(w*IN)), Emu(int(h*IN)))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp

def boxtext(sp, lines, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT):
    tf = sp.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = Emu(int(0.12*IN)); tf.margin_right = Emu(int(0.08*IN))
    tf.margin_top = Emu(int(0.05*IN)); tf.margin_bottom = Emu(int(0.05*IN))
    for i, (t, sz, b, c) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.alignment = align
        r = p.add_run(); r.text = t; r.font.size = Pt(sz); r.font.bold = b
        if c is not None: r.font.color.rgb = c
        _ea(r)

def pic(s, path, l, t, w=None, h=None):
    kw = {}
    if w: kw["width"] = Emu(int(w*IN))
    if h: kw["height"] = Emu(int(h*IN))
    if os.path.exists(path):
        s.shapes.add_picture(path, Emu(int(l*IN)), Emu(int(t*IN)), **kw)

FIG_FLOW = os.path.join(OUTDIR, "mechanism_flow.png")
FIG_ALDOL = os.path.join(OUTDIR, "aldol_scheme.png")
FIG_POLY = os.path.join(OUTDIR, "polyenal_long.png")
FIG_APHA = os.path.join(OUTDIR, "apha_trend.png")
FIG_UV = os.path.join(OUTDIR, "uv_spectrum.png")
FIG_UV450 = os.path.join(OUTDIR, "uv450_trend.png")

# ════════════════════════════════════════════════════════════
# 1. 表紙
# ════════════════════════════════════════════════════════════
s = newslide(LY_COVER)
remove_content_ph(s)
# 淡いアクセント帯＋大きめタイトル（明示配置で確実に表紙らしく）
rrect(s, 0.0, 2.45, 13.333, 0.06, BLUE)
rrect(s, 0.0, 4.55, 13.333, 0.04, RGBColor(0xC9, 0xD2, 0xD8))
tf = tbx(s, 0.9, 2.75, 11.5, 1.6)
P(tf, "製品DEG色相悪化　対応", 34, bold=True, color=DBLUE, first=True, space=4)
P(tf, "事実の整理 → 推定メカニズム → 対応", 17, color=GRAY, space=2)
tf2 = tbx(s, 0.9, 4.75, 11.5, 0.6)
P(tf2, M.DATE, 14, color=GRAY, first=True)
# 下部：3フェーズのロードマップ
road = [("① 事実の整理", "発生事象・着色物質・工程別グリッド"),
        ("② 推定メカニズム", "アルドール縮合→前駆体→タンク熟成→450nm"),
        ("③ 対応", "触媒交換・運転条件・UV450判定")]
bx_w = 3.8; gap = 0.3; xx = 0.9
for t1, t2 in road:
    bx = rrect(s, xx, 5.75, bx_w, 1.05, H_FILL, line=BLUE, lw=1.0)
    boxtext(bx, [(t1, 14, True, DBLUE), (t2, 9, False, GRAY)], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    xx += bx_w + gap

# ════════════════════════════════════════════════════════════
# 1.5 概要（結論サマリ：承認者が最初に全体を掴む）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "概要")
y = 1.55
tagfill = {"事実": C_OPE, "仮説": C_LIT, "対応": H_FILL, "ご依頼": C_SUJI}
for lab, txt in M.SUMMARY:
    bx = rrect(s, 0.6, y, 12.1, 1.18, tagfill.get(lab, H_FILL), line=LINE, lw=1.0)
    tag = rrect(s, 0.8, y + 0.32, 1.5, 0.54, WHITE, line=BLUE, lw=1.25)
    boxtext(tag, [(lab, 14, True, DBLUE)], align=PP_ALIGN.CENTER)
    tf = tbx(s, 2.55, y + 0.12, 10.0, 0.95)
    P(tf, txt, 13, color=INK, first=True)
    y += 1.32

# ════════════════════════════════════════════════════════════
# 2. 発生事象（事実）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "発生事象（事実）")
tf = tbx(s, 0.55, 1.5, 6.7, 5.4)
for i, t in enumerate(M.PHENOMENON):
    P(tf, t, 13.5, bullet=True, first=(i == 0), space=12)
pic(s, FIG_APHA, 7.5, 2.0, w=5.4)

# ════════════════════════════════════════════════════════════
# 3. 着色物質（事実・推定）＋ ポリエナール構造
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "着色物質と着色現象")
fb = rrect(s, 0.55, 1.4, 6.6, 2.35, C_OPE, line=LINE, lw=1.0)
tf = fb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
tf.margin_left = Emu(int(0.12*IN)); tf.margin_top = Emu(int(0.06*IN))
P(tf, "■ 確認した事実", 12.5, bold=True, color=DBLUE, first=True, space=2)
for t in M.COLORANT_FACT: P(tf, t, 11, bullet=True, space=2)
eb = rrect(s, 0.55, 3.9, 6.6, 3.0, C_LIT, line=LINE, lw=1.0)
tf = eb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
tf.margin_left = Emu(int(0.12*IN)); tf.margin_top = Emu(int(0.06*IN))
P(tf, "■ ここから読める推定", 12.5, bold=True, color=DBLUE, first=True, space=2)
for t in M.COLORANT_EST: P(tf, t, 11, bullet=True, space=2)
# 右：UVスペクトル＋ポリエナール構造
pic(s, FIG_UV, 7.35, 1.4, w=5.6)
pic(s, FIG_POLY, 7.6, 4.7, w=5.1)
tf2 = tbx(s, 7.4, 6.45, 5.6, 0.7)
P(tf2, "着色物質＝共役ポリエナール CH3-(CH=CH)n-CHO（数〜10ppbで発色）", 10.5, bold=True, color=DBLUE, first=True)

# ════════════════════════════════════════════════════════════
# 4. 推定メカニズム（化学反応スキーム）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, M.MECH_CAPTION)
pic(s, FIG_ALDOL, 0.5, 1.45, w=12.33)
tf = tbx(s, 0.6, 5.55, 12.1, 1.0)
P(tf, "一筋： " + M.MECH_ONELINE, 12.5, bold=True, color=DBLUE, first=True, space=3)
tf2 = tbx(s, 0.6, 6.75, 12.1, 0.6)
P(tf2, "※ " + M.MECH_CAVEAT, 10, color=GRAY, first=True)

# ════════════════════════════════════════════════════════════
# 5. 工程フロー（どこで何が起きるか）※工程別グリッドの詳細は事実整理Excelに集約
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "縮合は脱水塔以降で起こる（工程フロー）")
pic(s, FIG_FLOW, 0.35, 1.75, w=12.6)
tf = tbx(s, 0.6, 5.7, 12.1, 1.2)
P(tf, "工程ごとの事実（運転・解放・RD）／文献・社外知見／一般原理の切り分けは「事実整理（工程別グリッド）」に整理。",
  11, color=GRAY, first=True)

# ════════════════════════════════════════════════════════════
# 6.5 工程別サマリ（Excelグリッドの織り込み：工程×裏付け強度×小結論）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "工程別サマリ（事実の裏付け強度と小結論）")
tf = tbx(s, 0.6, 1.32, 12.1, 0.5)
P(tf, "裏付け強度 ◎事実複数で確定／○一部事実＋原理／△収支・原理のみで確度低。工程ごとの事実の詳細は「事実整理（工程別グリッド）」に集約。", 10.5, color=GRAY, first=True)
concl = M.GRID[-1][2]
t3 = s.shapes.add_table(len(M.STAGES)+1, 3, Emu(int(0.6*IN)), Emu(int(1.95*IN)), Emu(int(12.1*IN)), Emu(int(4.75*IN))).table
t3.columns[0].width = Emu(int(3.0*IN)); t3.columns[1].width = Emu(int(1.4*IN)); t3.columns[2].width = Emu(int(7.7*IN))
def cell3(i, j, text, bold=False, fill=WHITE, color=INK, size=11, align=None):
    c = t3.cell(i, j); c.text = ""; c.vertical_anchor = MSO_ANCHOR.MIDDLE
    c.margin_top = Emu(int(0.03*IN)); c.margin_bottom = Emu(int(0.03*IN)); c.margin_left = Emu(int(0.06*IN))
    p = c.text_frame.paragraphs[0]
    if align is not None: p.alignment = align
    r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
    c.fill.solid(); c.fill.fore_color.rgb = fill
cell3(0, 0, "工程", bold=True, color=DBLUE, fill=H_FILL)
cell3(0, 1, "裏付け", bold=True, color=DBLUE, fill=H_FILL, align=PP_ALIGN.CENTER)
cell3(0, 2, "この工程で言えること", bold=True, color=DBLUE, fill=H_FILL)
symcol = {"◎": BLUE, "○": DBLUE, "△": RGBColor(0xC0, 0x7A, 0x00)}
for i, (stg, sym, c6) in enumerate(zip(M.STAGES, M.STRENGTH, concl), 1):
    cell3(i, 0, stg.replace("\n", " "), bold=True, color=DBLUE, fill=PROC, size=10.5)
    cell3(i, 1, sym, bold=True, color=symcol.get(sym, DBLUE), fill=WHITE, size=15, align=PP_ALIGN.CENTER)
    cell3(i, 2, c6, size=10)

# ════════════════════════════════════════════════════════════
# 7. なぜ今回（条件の重なり）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "なぜ今回だけ起きたか（条件の重なり）")
tf = tbx(s, 0.6, 1.45, 12.1, 0.5)
P(tf, "過去にもアルデヒドが高い時期はあったが、今回は反応させてしまう条件がそろった。", 12.5, color=GRAY, first=True)
y = 2.05
for head, txt, act in M.WHY_NOW:
    bx = rrect(s, 0.6, y, 12.1, 1.05, H_FILL, line=LINE, lw=1.0)
    tf = bx.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(int(0.14*IN))
    p = tf.paragraphs[0]; r = p.add_run(); r.text = "【" + head + "】 "; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = DBLUE; _ea(r)
    r2 = p.add_run(); r2.text = txt; r2.font.size = Pt(11.5); r2.font.color.rgb = INK; _ea(r2)
    p2 = tf.add_paragraph(); r3 = p2.add_run(); r3.text = act; r3.font.size = Pt(11.5); r3.font.bold = True; r3.font.color.rgb = BLUE; _ea(r3)
    y += 1.16

# ════════════════════════════════════════════════════════════
# 8. 設備（鉄サビ）は主因でない
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "設備（鉄サビ）は色相の主因ではない")
pic(s, FIG_UV, 6.7, 1.55, w=6.3)
tf = tbx(s, 0.55, 1.7, 5.95, 4.0)
_pts = [
    "C-1503の黒色付着物はほぼ鉄サビ＋有機酸（局所の腐食）。",
    "付着物をDEGに加えても380nmのみ＝実機の450nmは再現しない（保管試験）。",
    "T-555残液はAPHA30で頭打ち＝タンクの鉄サビは促進触媒でない。",
]
for i, t in enumerate(_pts):
    P(tf, t, 13, bullet=True, first=(i == 0), space=16)
concl = rrect(s, 0.55, 5.75, 12.2, 0.95, C_SUJI, line=BLUE, lw=1.0)
boxtext(concl, [("→ 鉄サビ・付着物は色相（450nm）の主因でなく、寄与は局所380nm。実測で判断（金属が無くてもアルデヒドとpHで着色を説明できる）。洗浄は予防保全。", 12, True, DBLUE)])

# ════════════════════════════════════════════════════════════
# 9. 開放・検査（機器ごと）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "設備に色相の主因なし（開放・検査の結論）")
tf = tbx(s, 0.6, 1.4, 12.1, 0.55); P(tf, M.INSPECTION_LOGIC, 11, color=GRAY, first=True)
rows = M.INSPECTION
t2 = s.shapes.add_table(len(rows)+1, 2, Emu(int(0.6*IN)), Emu(int(2.05*IN)), Emu(int(12.1*IN)), Emu(int(4.5*IN))).table
t2.columns[0].width = Emu(int(3.6*IN)); t2.columns[1].width = Emu(int(8.5*IN))
def cell2(i, j, text, bold=False, fill=WHITE, color=INK):
    c = t2.cell(i, j); c.text = ""; c.vertical_anchor = MSO_ANCHOR.MIDDLE
    c.margin_top = Emu(int(0.03*IN)); c.margin_bottom = Emu(int(0.03*IN)); c.margin_left = Emu(int(0.06*IN))
    p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = text
    r.font.size = Pt(10.5); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
    c.fill.solid(); c.fill.fore_color.rgb = fill
cell2(0, 0, "機器", bold=True, color=DBLUE, fill=H_FILL); cell2(0, 1, "開放理由・結果", bold=True, color=DBLUE, fill=H_FILL)
for i, (eq, txt) in enumerate(rows, 1):
    cell2(i, 0, eq, bold=True, color=DBLUE, fill=PROC); cell2(i, 1, txt)

# ════════════════════════════════════════════════════════════
# 10. 対応（運転条件）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "対応（再発防止・スタートアップ運転条件）")
rows = M.COUNTERMEASURE
t4 = s.shapes.add_table(len(rows)+1, 3, Emu(int(0.6*IN)), Emu(int(1.5*IN)), Emu(int(12.1*IN)), Emu(int(5.0*IN))).table
t4.columns[0].width = Emu(int(2.6*IN)); t4.columns[1].width = Emu(int(7.7*IN)); t4.columns[2].width = Emu(int(1.8*IN))
def cell4(i, j, text, bold=False, fill=WHITE, color=INK, size=11, align=None):
    c = t4.cell(i, j); c.text = ""; c.vertical_anchor = MSO_ANCHOR.MIDDLE
    c.margin_top = Emu(int(0.03*IN)); c.margin_bottom = Emu(int(0.03*IN)); c.margin_left = Emu(int(0.06*IN))
    p = c.text_frame.paragraphs[0]
    if align is not None: p.alignment = align
    r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
    c.fill.solid(); c.fill.fore_color.rgb = fill
cell4(0, 0, "項目", bold=True, color=DBLUE, fill=H_FILL)
cell4(0, 1, "内容", bold=True, color=DBLUE, fill=H_FILL)
cell4(0, 2, "状況", bold=True, color=DBLUE, fill=H_FILL, align=PP_ALIGN.CENTER)
for i, (head, txt, conf) in enumerate(rows, 1):
    cell4(i, 0, head, bold=True, color=DBLUE, fill=PROC, size=10.5)
    cell4(i, 1, txt, size=9.5)
    cell4(i, 2, conf, size=9.5, align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════
# 11. 品質管理・早期判定
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "出荷再開に向けた品質管理・早期判定")
tf = tbx(s, 0.55, 1.5, 7.1, 5.4)
for i, t in enumerate(M.QC):
    P(tf, t, 12, bullet=True, first=(i == 0), space=11)
pic(s, FIG_UV450, 7.8, 2.5, w=5.3)

# ════════════════════════════════════════════════════════════
# 11.5 出荷判定の運用（歯止め）：枠組みを明示（具体値は会議で確定）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "出荷判定の運用（歯止め）")
tf = tbx(s, 0.6, 1.35, 12.1, 0.5)
P(tf, M.GATE_NOTE, 11, color=GRAY, first=True)
rows = M.GATE
tg = s.shapes.add_table(len(rows)+1, 2, Emu(int(0.6*IN)), Emu(int(1.95*IN)), Emu(int(12.1*IN)), Emu(int(4.7*IN))).table
tg.columns[0].width = Emu(int(3.2*IN)); tg.columns[1].width = Emu(int(8.9*IN))
def cellg(i, j, text, bold=False, fill=WHITE, color=INK, size=11):
    c = tg.cell(i, j); c.text = ""; c.vertical_anchor = MSO_ANCHOR.MIDDLE
    c.margin_top = Emu(int(0.03*IN)); c.margin_bottom = Emu(int(0.03*IN)); c.margin_left = Emu(int(0.06*IN))
    p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
    c.fill.solid(); c.fill.fore_color.rgb = fill
cellg(0, 0, "項目", bold=True, color=DBLUE, fill=H_FILL)
cellg(0, 1, "運用（「本会議で確定」＝具体値は本対策会議で固める）", bold=True, color=DBLUE, fill=H_FILL)
for i, (k, v) in enumerate(rows, 1):
    cellg(i, 0, k, bold=True, color=DBLUE, fill=PROC, size=10.5)
    cellg(i, 1, v, size=10)

# ════════════════════════════════════════════════════════════
# 12. 承認のお願い
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "出荷再開の判断（ご承認のお願い）")
steps = [
    ("① まず出口実測で歯止め", "初期流動品は、全項目が規格内、かつタンク保管中のUV・色相が経時で不変であることを確認したうえでのみ出荷する。仮説の真偽に依らず、実測で歯止めをかける。"),
    ("② それを支える対応", "触媒交換（実施済）・設備洗浄・運転条件（NaOH非投入／温度を下げない／リン酸で緩和）。早期判定はUV450nm主・UV330補助。判定基準の具体値は本会議で確定。"),
    ("③ なぜ起きたか（参考）", "触媒劣化で増えたアルデヒドが縮合して前駆体になり、タンクで熟成・共役伸長して450nm（最有力仮説・ベンチ未再現）。"),
]
y = 1.5
for head, txt in steps:
    bx = rrect(s, 0.6, y, 12.1, 1.18, H_FILL, line=LINE, lw=1.0)
    tf = bx.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(int(0.16*IN)); tf.margin_right = Emu(int(0.12*IN))
    p = tf.paragraphs[0]; r = p.add_run(); r.text = head; r.font.size = Pt(13.5); r.font.bold = True; r.font.color.rgb = DBLUE; _ea(r)
    p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = txt; r2.font.size = Pt(11.5); r2.font.color.rgb = INK; _ea(r2)
    y += 1.3
bar = rrect(s, 0.6, 5.55, 12.1, 0.95, C_SUJI, line=BLUE, lw=1.25)
boxtext(bar, [("上記の前提（出口実測での歯止め）で、製品DEGの出荷再開についてご承認をお願いします。", 15, True, DBLUE)], align=PP_ALIGN.CENTER)

# ── 元テンプレートの既存8枚を除去（新規スライドだけ残す） ──
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
# 新規スライドは末尾。元の先頭8枚を除去。
n_orig = 8
kept_rids = set()
for e in ids[n_orig:]:
    kept_rids.add(e.get(qn("r:id")))
for e in ids[:n_orig]:
    sldIdLst.remove(e)
# orphan rels を落とす
for rId, rel in list(prs.part.rels.items()):
    if rel.reltype.endswith("/slide") and rId not in kept_rids:
        try: prs.part.drop_rel(rId)
        except Exception: pass

os.makedirs(OUTDIR, exist_ok=True)
prs.save(OUT)
# dup check
import zipfile, collections
z = zipfile.ZipFile(OUT); names = z.namelist()
dup = [k for k, v in collections.Counter(names).items() if v > 1]
print("saved:", os.path.normpath(OUT), "| slides:", len(prs.slides._sldIdLst), "| dup:", dup or "なし")
