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

# ════════════════════════════════════════════════════════════
# 1. 表紙
# ════════════════════════════════════════════════════════════
s = newslide(LY_COVER)
set_title(s, "製品DEG色相悪化　対応")
for ph in s.placeholders:
    if ph.placeholder_format.idx not in (0,):
        try:
            if ph.has_text_frame and ph.placeholder_format.idx in (1,):
                ph.text = ""; r = ph.text_frame.paragraphs[0].add_run()
                r.text = "事実の整理 → 推定メカニズム → 対応"; r.font.size = Pt(15); r.font.color.rgb = GRAY; _ea(r)
        except Exception:
            pass

# ════════════════════════════════════════════════════════════
# 2. 発生事象（事実）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "発生事象（事実）")
tf = tbx(s, 0.6, 1.5, 12.1, 5.4)
for i, t in enumerate(M.PHENOMENON):
    P(tf, t, 15, bullet=True, first=(i == 0), space=10)

# ════════════════════════════════════════════════════════════
# 3. 着色物質（事実・推定）＋ ポリエナール構造
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "着色物質と着色現象")
fb = rrect(s, 0.55, 1.45, 7.4, 2.3, C_OPE, line=LINE, lw=1.0)
tf = fb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
tf.margin_left = Emu(int(0.12*IN)); tf.margin_top = Emu(int(0.06*IN))
P(tf, "■ 確認した事実", 12.5, bold=True, color=DBLUE, first=True, space=2)
for t in M.COLORANT_FACT: P(tf, t, 11, bullet=True, space=2)
eb = rrect(s, 0.55, 3.95, 7.4, 2.95, C_LIT, line=LINE, lw=1.0)
tf = eb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
tf.margin_left = Emu(int(0.12*IN)); tf.margin_top = Emu(int(0.06*IN))
P(tf, "■ ここから読める推定", 12.5, bold=True, color=DBLUE, first=True, space=2)
for t in M.COLORANT_EST: P(tf, t, 11, bullet=True, space=2)
# 構造（ポリエナール）
if os.path.exists(FIG_POLY):
    pic(s, FIG_POLY, 8.25, 1.7, w=4.5)
tf2 = tbx(s, 8.25, 4.9, 4.5, 2.0)
P(tf2, "着色物質＝共役ポリエナール", 12, bold=True, color=DBLUE, first=True, space=2)
P(tf2, "CH3-(CH=CH)n-CHO", 12, bold=True, color=INK, space=2)
P(tf2, "共役が伸びるほど長波長化し、十分伸びると450nm（黄）。数〜10ppbの微量で発色。", 10.5, color=GRAY, space=2)

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
# 5. 工程別の整理（①運転 ②RD ③文献 → ④筋）＝表
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "工程別の整理（①運転 ②RD ③文献 → 筋）")
nrow = len(M.PROCESS_MATRIX) + 2  # header + rows + suji
tbl = s.shapes.add_table(nrow, 5, Emu(int(0.4*IN)), Emu(int(1.4*IN)), Emu(int(12.5*IN)), Emu(int(5.5*IN))).table
ws = [1.7, 2.9, 2.9, 2.9, 2.1]
for j, w in enumerate(ws): tbl.columns[j].width = Emu(int(w*IN))
hdr = M.MATRIX_COLS
hfills = [PROC, C_OPE, C_RD, C_LIT, C_SUJI]
def cell(i, j, text, bold=False, color=INK, fill=WHITE, size=9):
    c = tbl.cell(i, j); c.text = ""; c.margin_left = Emu(int(0.05*IN)); c.margin_right = Emu(int(0.03*IN))
    c.margin_top = Emu(int(0.02*IN)); c.margin_bottom = Emu(int(0.02*IN))
    c.vertical_anchor = MSO_ANCHOR.TOP
    p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
    c.fill.solid(); c.fill.fore_color.rgb = fill
for j, h in enumerate(hdr): cell(0, j, h, bold=True, color=DBLUE, fill=hfills[j], size=9.5)
for i, (proc, ope, rd, lit, suji) in enumerate(M.PROCESS_MATRIX, 1):
    cell(i, 0, proc, bold=True, color=DBLUE, fill=PROC, size=9)
    cell(i, 1, ope, fill=C_OPE); cell(i, 2, rd, fill=C_RD); cell(i, 3, lit, fill=C_LIT)
    cell(i, 4, suji, bold=True, color=DBLUE, fill=C_SUJI)
# suji row
li = nrow - 1
cell(li, 0, "トータル＝1本の筋", bold=True, color=WHITE, fill=BLUE, size=9)
c = tbl.cell(li, 1);
tbl.cell(li, 1).merge(tbl.cell(li, 4))
cc = tbl.cell(li, 1); cc.text = ""; cc.fill.solid(); cc.fill.fore_color.rgb = C_SUJI
p = cc.text_frame.paragraphs[0]; r = p.add_run(); r.text = M.MECH_ONELINE
r.font.size = Pt(9.5); r.font.bold = True; r.font.color.rgb = DBLUE; _ea(r)

# ════════════════════════════════════════════════════════════
# 6. 工程フロー（どこで何が起きるか）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "どこで何が起きるか（工程フロー）")
pic(s, FIG_FLOW, 0.4, 1.7, w=12.5)

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
tf = tbx(s, 0.6, 1.5, 12.1, 5.4)
for i, t in enumerate(M.RUST):
    last = (i == len(M.RUST) - 1)
    P(tf, t, 13.5, bullet=True, first=(i == 0), bold=last, color=(DBLUE if last else INK), space=11)

# ════════════════════════════════════════════════════════════
# 9. 開放・検査（機器ごと）
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "開放・検査の整理（仮説駆動）")
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
tf = tbx(s, 0.6, 1.5, 12.1, 5.4)
for head, txt, conf in M.COUNTERMEASURE:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = "【" + head + "】 "; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = DBLUE; _ea(r)
    r2 = p.add_run(); r2.text = txt + " "; r2.font.size = Pt(12); r2.font.color.rgb = INK; _ea(r2)
    r3 = p.add_run(); r3.text = "（" + conf + "）"; r3.font.size = Pt(11.5); r3.font.bold = True; r3.font.color.rgb = BLUE; _ea(r3)
    p.space_after = Pt(8)

# ════════════════════════════════════════════════════════════
# 11. 品質管理・早期判定
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "出荷再開に向けた品質管理・早期判定")
tf = tbx(s, 0.6, 1.5, 12.1, 5.4)
for i, t in enumerate(M.QC):
    P(tf, t, 13, bullet=True, first=(i == 0), space=10)

# ════════════════════════════════════════════════════════════
# 12. 承認のお願い
# ════════════════════════════════════════════════════════════
s = newslide(LY_TEXT); remove_content_ph(s); set_title(s, "出荷再開の判断（ご承認のお願い）")
tf = tbx(s, 0.6, 1.5, 12.1, 3.6)
P(tf, "触媒交換（原因除去）・設備洗浄・スタートアップ運転条件（NaOH非投入ベース／温度を下げない／リン酸で緩和）により、色相悪化の再発を抑える。", 14, first=True, space=12)
P(tf, "再稼働後の初期流動品は、全項目規格内かつタンク保管中の色相・UVの経時変化が無いことを確認したうえで出荷する。", 14, space=12)
P(tf, "早期判定はUV450nm（最も感度が高い）を主指標とし、UV330nmを補助に用いる。判定基準の具体値は本対策会議で確定する。", 14, space=12)
bar = rrect(s, 0.6, 5.45, 12.1, 0.95, H_FILL, line=BLUE, lw=1.25)
boxtext(bar, [("上記の前提で、製品DEGの出荷再開についてご承認をお願いします。", 15, True, DBLUE)], align=PP_ALIGN.CENTER)

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
