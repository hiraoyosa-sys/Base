# -*- coding: utf-8 -*-
"""対策会議用PPTX：製品DEG色相悪化（事実→仮説→対応）。配布クリーン。"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import deg_master as M

OUT = os.path.join(os.path.dirname(__file__), "..", "outputs", "製品DEG色相悪化_対策会議.pptx")
JP = "游ゴシック"
IN = 914400
NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x5B, 0xA8)
SUB = RGBColor(0xD9, 0xE1, 0xF2)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC0, 0x39, 0x2B)
GRAY = RGBColor(0x55, 0x55, 0x55)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xEA, 0xEF, 0xF7)

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
cp = prs.core_properties
cp.author = ""; cp.last_modified_by = ""; cp.title = "製品DEG色相悪化 対策会議"; cp.comments = ""
BLANK = prs.slide_layouts[6]

def slide():
    return prs.slides.add_slide(BLANK)

def font(run, size, bold=False, color=None):
    run.font.name = JP; run.font.size = Pt(size); run.font.bold = bold
    run.font._rPr.set("{http://schemas.openxmlformats.org/drawingml/2006/main}", "")  # noop guard
    if color is not None: run.font.color.rgb = color
    ea = run.font._rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
    return run

def _ea(run):
    rPr = run.font._rPr
    from pptx.oxml.ns import qn
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", JP)

def box(s, x, y, w, h):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    return tf

def put(tf, text, size, bold=False, color=None, first=False, bullet=False, space=4):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if bullet: text = "・" + text
    r = p.add_run(); r.text = text; r.font.name = JP; r.font.size = Pt(size); r.font.bold = bold
    if color is not None: r.font.color.rgb = color
    _ea(r); p.space_after = Pt(space)
    return p

def rect(s, x, y, w, h, fill, line=None, lw=1.0):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp

def boxtext(sp, lines, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT):
    tf = sp.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = Emu(int(0.12*IN)); tf.margin_right = Emu(int(0.08*IN))
    tf.margin_top = Emu(int(0.05*IN)); tf.margin_bottom = Emu(int(0.05*IN))
    for i, (t, sz, b, c) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.alignment = align
        r = p.add_run(); r.text = t; r.font.name = JP; r.font.size = Pt(sz); r.font.bold = b
        if c is not None: r.font.color.rgb = c
        _ea(r)

def header(s, text, step=None):
    bar = rect(s, 0, 0, 13.333, 0.92, NAVY)
    boxtext(bar, [(text, 23, True, WHITE)], anchor=MSO_ANCHOR.MIDDLE)
    if step:
        tag = rect(s, 11.4, 0.22, 1.7, 0.48, BLUE)
        boxtext(tag, [(step, 12, True, WHITE)], align=PP_ALIGN.CENTER)

def body(s):
    return box(s, 0.55, 1.15, 12.25, 5.9)

# ── 1. タイトル ──
s = slide()
rect(s, 0, 0, 13.333, 7.5, NAVY)
rect(s, 0, 2.5, 13.333, 2.5, RGBColor(0x16, 0x2A, 0x4A))
tf = box(s, 0.9, 2.7, 11.5, 2.1)
put(tf, "製品DEG色相悪化　対策", 34, True, WHITE, first=True)
put(tf, "事実の整理 → 推定メカニズム → 対応", 18, False, SUB, space=2)
tf2 = box(s, 0.9, 6.4, 11.5, 0.8)
put(tf2, M.DATE, 14, False, SUB, first=True)

# ── 2. 発生事象 ──
s = slide(); header(s, "発生事象（事実）", "事実")
tf = body(s)
for i, t in enumerate(M.PHENOMENON):
    put(tf, t, 15.5, bullet=True, first=(i == 0), space=8)

# ── 3. 着色物質と現象 ──
s = slide(); header(s, "着色物質と着色現象", "事実")
fb = rect(s, 0.55, 1.1, 12.25, 2.75, RGBColor(0xE2, 0xEF, 0xDA), line=GREEN, lw=1.25)
tf = fb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
tf.margin_left = Emu(int(0.14*IN)); tf.margin_top = Emu(int(0.08*IN))
p = tf.paragraphs[0]; r = p.add_run(); r.text = "■ 確認した事実"; r.font.name = JP; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = GREEN; _ea(r)
for t in M.COLORANT_FACT:
    pp = tf.add_paragraph(); rr = pp.add_run(); rr.text = "・" + t; rr.font.name = JP; rr.font.size = Pt(12); _ea(rr); pp.space_after = Pt(2)
eb = rect(s, 0.55, 4.0, 12.25, 2.75, LIGHT, line=GRAY, lw=1.0)
tf = eb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
tf.margin_left = Emu(int(0.14*IN)); tf.margin_top = Emu(int(0.08*IN))
p = tf.paragraphs[0]; r = p.add_run(); r.text = "■ ここから読める推定"; r.font.name = JP; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = NAVY; _ea(r)
for t in M.COLORANT_EST:
    pp = tf.add_paragraph(); rr = pp.add_run(); rr.text = "・" + t; rr.font.name = JP; rr.font.size = Pt(12); _ea(rr); pp.space_after = Pt(2)

# ── 4. 推定メカニズム（フロー図） ──
s = slide(); header(s, M.MECH_CAPTION, "仮説")
import os as _os
_fig = _os.path.join(_os.path.dirname(__file__), "..", "outputs", "mechanism_flow.png")
if _os.path.exists(_fig):
    s.shapes.add_picture(_fig, Inches(0.5), Inches(1.15), width=Inches(12.33))
tf = box(s, 0.55, 4.6, 12.25, 1.2)
put(tf, "一筋： " + M.MECH_ONELINE, 13.5, bold=True, color=NAVY, first=True, space=4)
tf2 = box(s, 0.55, 6.15, 12.25, 1.0)
put(tf2, "※ " + M.MECH_CAVEAT, 11, False, GRAY, first=True)

# ── 5. なぜ今回 ──
s = slide(); header(s, "なぜ今回だけ起きたか（条件の重なり）", "仮説")
tf = body(s)
put(tf, "過去にもアルデヒドが高い時期はあったが、今回は反応させてしまう条件がそろった。", 14, False, GRAY, first=True, space=10)
for head, txt, act in M.WHY_NOW:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = "【" + head + "】 "; r.font.name = JP; r.font.size = Pt(14.5); r.font.bold = True; r.font.color.rgb = NAVY; _ea(r)
    r2 = p.add_run(); r2.text = txt + " "; r2.font.name = JP; r2.font.size = Pt(13.5); _ea(r2)
    r3 = p.add_run(); r3.text = act; r3.font.name = JP; r3.font.size = Pt(13.5); r3.font.bold = True; r3.font.color.rgb = GREEN; _ea(r3)
    p.space_after = Pt(9)

# ── 6. 設備の関与の評価 ──
s = slide(); header(s, "設備（鉄サビ）は色相の主因ではない", "事実")
tf = body(s)
for i, t in enumerate(M.RUST):
    col = GREEN if i == len(M.RUST) - 1 else None
    bold = (i == len(M.RUST) - 1)
    put(tf, t, 14.5, bold=bold, color=col, bullet=True, first=(i == 0), space=9)

# ── 7. 開放・検査 ──
s = slide(); header(s, "開放・検査の整理（仮説駆動）", "事実")
tf = box(s, 0.55, 1.1, 12.25, 0.7)
put(tf, M.INSPECTION_LOGIC, 12.5, False, GRAY, first=True)
# テーブル
rows = M.INSPECTION
tbl = s.shapes.add_table(len(rows)+1, 2, Inches(0.55), Inches(1.95), Inches(12.25), Inches(4.6)).table
tbl.columns[0].width = Inches(3.6); tbl.columns[1].width = Inches(8.65)
hd = ["機器", "開放理由・結果"]
for j, htext in enumerate(hd):
    c = tbl.cell(0, j); c.text = ""
    p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = htext
    r.font.name = JP; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = WHITE; _ea(r)
    c.fill.solid(); c.fill.fore_color.rgb = NAVY
for i, (eq, txt) in enumerate(rows, 1):
    for j, val in enumerate((eq, txt)):
        c = tbl.cell(i, j); c.text = ""
        p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = val
        r.font.name = JP; r.font.size = Pt(11); r.font.bold = (j == 0); _ea(r)
        c.fill.solid(); c.fill.fore_color.rgb = SUB if j == 0 else WHITE

# ── 8. 対応 ──
s = slide(); header(s, "対応（再発防止・スタートアップ運転条件）", "対応")
tf = body(s)
for head, txt, conf in M.COUNTERMEASURE:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = "【" + head + "】 "; r.font.name = JP; r.font.size = Pt(13.5); r.font.bold = True; r.font.color.rgb = NAVY; _ea(r)
    r2 = p.add_run(); r2.text = txt + " "; r2.font.name = JP; r2.font.size = Pt(12.5); _ea(r2)
    r3 = p.add_run(); r3.text = "（" + conf + "）"; r3.font.name = JP; r3.font.size = Pt(12); r3.font.bold = True; r3.font.color.rgb = BLUE; _ea(r3)
    p.space_after = Pt(7)

# ── 9. 品質管理・早期判定 ──
s = slide(); header(s, "出荷再開に向けた品質管理・早期判定", "対応")
tf = body(s)
for i, t in enumerate(M.QC):
    put(tf, t, 14.5, bullet=True, first=(i == 0), space=9)

# ── 10. 残論点 ──
s = slide(); header(s, "残論点（要確認）", "対応")
tf = body(s)
for head, txt in M.OPEN_ISSUES:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = "【" + head + "】 "; r.font.name = JP; r.font.size = Pt(14.5); r.font.bold = True; r.font.color.rgb = NAVY; _ea(r)
    r2 = p.add_run(); r2.text = txt; r2.font.name = JP; r2.font.size = Pt(13.5); _ea(r2)
    p.space_after = Pt(9)

# ── 11. 承認のお願い ──
s = slide(); header(s, "出荷再開の判断（ご承認のお願い）", "対応")
tf = body(s)
put(tf, "触媒交換（原因除去）・設備洗浄・スタートアップ運転条件（NaOH非投入ベース／温度を下げない／リン酸で緩和）により、色相悪化の再発を抑える。", 15.5, first=True, space=12)
put(tf, "再稼働後の初期流動品は、全項目規格内かつタンク保管中の色相・UVの経時変化が無いことを確認したうえで出荷する。", 15.5, space=12)
put(tf, "早期判定はUV450nm（最も感度が高い）を主指標とし、UV330nmを補助に用いる。", 15.5, space=12)
bar = rect(s, 0.55, 5.5, 12.25, 0.95, NAVY)
boxtext(bar, [("上記の前提で、製品DEGの出荷再開についてご承認をお願いします。", 16, True, WHITE)], align=PP_ALIGN.CENTER)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)
print("saved:", os.path.normpath(OUT), "| slides:", len(prs.slides._sldIdLst))
