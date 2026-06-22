# -*- coding: utf-8 -*-
"""まとめ資料(65f0890b)を直接修正：S6設備/S11基礎研の空欄記入、
『どこで・どういう反応』統合スライド追加、補足索引/用意資料スライド追加。新規作成でなく当該ファイルを編集。"""
import os
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER

SRC = "/root/.claude/uploads/0ac036e2-7cd8-5317-a99c-aa39021935be/65f0890b-__DEG________EO__.pptx"
OUT = os.path.join(os.path.dirname(__file__), "..", "outputs", "DEG色相対応_SU前会議.pptx")
IN = 914400
EA = "游ゴシック"
BLUE = RGBColor(0x00, 0x5B, 0xAB); DBLUE = RGBColor(0x00, 0x3F, 0x7E); BLACK = RGBColor(0x00, 0x00, 0x00)
INK = RGBColor(0x22, 0x22, 0x22); GRAY = RGBColor(0x66, 0x66, 0x66); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
H_FILL = RGBColor(0xEB, 0xEF, 0xF2); PROC = RGBColor(0xD7, 0xE0, 0xE5); AMBER = RGBColor(0xFF, 0xF6, 0xE0)
LINE = RGBColor(0xC9, 0xD2, 0xD8)

prs = Presentation(SRC)
W = prs.slide_width / IN

def _ea(run):
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None: el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", EA)

def tbx(s, l, t, w, h):
    x = s.shapes.add_textbox(Emu(int(l*IN)), Emu(int(t*IN)), Emu(int(w*IN)), Emu(int(h*IN)))
    x.text_frame.word_wrap = True; return x.text_frame

def P(tf, text, size, bold=False, first=False, color=INK, bullet=False, space=5, align=None):
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    if align is not None: p.alignment = align
    r = p.add_run(); r.text = ("・" + text) if bullet else text
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
    p.space_after = Pt(space); return p

def rrect(s, l, t, w, h, fill, line=None, lw=1.0):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(int(l*IN)), Emu(int(t*IN)), Emu(int(w*IN)), Emu(int(h*IN)))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False; return sp

def arrow(s, l, t, w, h, color=BLUE):
    sp = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(int(l*IN)), Emu(int(t*IN)), Emu(int(w*IN)), Emu(int(h*IN)))
    sp.fill.solid(); sp.fill.fore_color.rgb = color; sp.line.fill.background(); sp.shadow.inherit = False; return sp

def vbox(s, l, t, w, h, lines, fill, line=BLUE, lw=1.1):
    sp = rrect(s, l, t, w, h, fill, line=line, lw=lw)
    tf = sp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(int(0.05*IN)); tf.margin_right = Emu(int(0.05*IN)); tf.margin_top = Emu(int(0.03*IN)); tf.margin_bottom = Emu(int(0.03*IN))
    for i, (t2, sz, b, c) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = t2; r.font.size = Pt(sz); r.font.bold = b; r.font.color.rgb = c; _ea(r)
    return sp

def title_of(s):
    for ph in s.placeholders:
        if ph.placeholder_format.idx == 0: return ph.text_frame.text.strip()
    return ""

def slide_by_title(key):
    for s in prs.slides:
        if key in title_of(s): return s
    return None

def body_textbox(s, l, t, w, h, bullets, size=14):
    tf = tbx(s, l, t, w, h)
    for i, b in enumerate(bullets):
        if isinstance(b, tuple):
            txt, bold, col = b
            P(tf, txt, size, bold=bold, color=col, first=(i == 0), bullet=False, space=8)
        else:
            P(tf, b, size, bullet=True, first=(i == 0), space=8)
    return tf

# ── S6 設備開放で確認された事実（空欄を記入） ──
s6 = slide_by_title("設備開放で確認された事実")
if s6:
    body_textbox(s6, 0.7, 1.55, 12.0, 5.4, [
        ("■ 推定メカニズム上、縮合は水が抜ける脱水塔以降で起こると考えられるため、EG精製系（脱水塔C-1404以降）を中心に開放・検査した。", True, DBLUE),
        "EG精製系は塔・配管・タンクを開放。C-1503（DEG塔）トップの黒色付着物（ほぼ鉄サビ）以外に、異常な汚れ・活性金属・腐食は認められない（C-1501/C-1502はCS→SUS改造部を含め錆込みなし）。",
        "唯一の異物（C-1503付着物・T-555/T-615内異物）は鉄サビで、DEGに加えても450nm（本件の着色）を作らないと基礎研で確認＝色相悪化の原因ではない。",
        "各機器の肉厚は定修前の運転中に測定。原肉に対し有意な減肉なし。",
        ("→ 設備・金属が原因で色相が悪化したのではないと判断。残る上流側はメカニズム上、縮合の場ではないため追加開放は不要。", True, DBLUE),
    ], size=14)

# ── S11 基礎研で検証された事実（空欄を記入） ──
s11 = slide_by_title("基礎研で検証された事実")
if s11:
    body_textbox(s11, 0.7, 1.55, 12.0, 5.4, [
        ("■ 着色物質の正体", True, DBLUE),
        "GPCで分子量200〜800のブロード、GCの複数検体が457nmを捕捉＝共役ポリエナール類。存在量はppbオーダー（吸光度からの換算）。",
        ("■ 鉄サビ・付着物の寄与", True, DBLUE),
        "鉄サビ・C-1503付着物をDEGに添加しても380nmのみで450nmは出ない＝鉄サビは本件の着色を作らない。",
        ("■ 反応の条件", True, DBLUE),
        "酸素があると前駆体が有機酸へ酸化され縮合が進まない（窒素で進み、大気接触で止まる、と整合）。",
        "アルデヒド（A/F-ALD）を添加しただけではDEGの450nmは再現しない＝アルデヒド単独でなく、低水分・塩基などの条件併存が必要。ラボでEG反応器出口を模擬した添加ではUVが悪化（共存でより大）。",
        ("※ 定量的な再現は未達（要追検証）。", False, GRAY),
    ], size=13)

# ── 統合スライド『どこで・どういう反応が起こるか』（スライド5,6統合）を追加 ──
ly = None
for l in prs.slide_layouts:
    if l.name == "タイトルとコンテンツ": ly = l
if ly is None: ly = prs.slide_layouts[2]
s = prs.slides.add_slide(ly)
for ph in list(s.placeholders):
    if ph.placeholder_format.idx != 0 and ph.placeholder_format.type in (
        PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.CENTER_TITLE):
        ph._element.getparent().remove(ph._element)
for ph in s.placeholders:
    if ph.placeholder_format.idx == 0:
        ph.text = ""; r = ph.text_frame.paragraphs[0].add_run(); r.text = "どこで・どういう反応が起こるか（工程と反応）"
        r.font.size = Pt(24); r.font.bold = True; r.font.color.rgb = BLACK; _ea(r)
# 上段：工程フロー（7段）
stages = ["EO反応系", "EG反応系", "EG濃縮系", "EG脱水系", "MEG精製系", "DEG精製系", "製品タンク"]
n = len(stages); bw = 1.5; gap = (W - 0.6 - n*bw) / (n - 1); x = 0.3; y = 1.55; bh = 0.85
cx = []
for k, nm in enumerate(stages):
    last = (k == n-1)
    vbox(s, x, y, bw, bh, [(nm, 11, True, (RGBColor(0xC0,0x7A,0x00) if last else DBLUE))],
         (AMBER if last else H_FILL), line=(RGBColor(0xC0,0x7A,0x00) if last else BLUE))
    cx.append(x + bw); x += bw + gap
for k in range(n-1):
    arrow(s, cx[k]+0.02, y+bh/2-0.1, gap-0.04, 0.2)
# 帯：水の有無
band = rrect(s, 0.3, 2.5, W-0.6, 0.42, RGBColor(0xF5, 0xF8, 0xFC), line=LINE, lw=0.75)
tfb = band.text_frame; tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tfb.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
r = p.add_run(); r.text = "　水が多い（縮合は進まず、アルデヒドは遊離のまま）　｜　水が抜ける（脱水塔以降で縮合開始）　→　タンクで伸長"
r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = DBLUE; _ea(r)
# 下段：反応（分子）の進行
mol = [
    [("遊離アルデヒド", 11, True, DBLUE), ("CH3CHO", 12, False, BLACK), ("上流（水あり）", 9.5, False, GRAY)],
    [("アルドール縮合", 11, True, DBLUE), ("→ 脱水(−H2O)", 11, False, BLACK), ("脱水塔以降", 9.5, False, GRAY)],
    [("前駆体", 11, True, DBLUE), ("CH3-(CH=CH)3-CHO", 11, False, BLACK), ("共役4・UV330nm", 9.5, False, GRAY)],
    [("ポリエナール", 11, True, RGBColor(0xC0,0x7A,0x00)), ("CH3-(CH=CH)n-CHO", 11, False, BLACK), ("共役〜9・450nm（黄）", 9.5, True, RGBColor(0xC0,0x7A,0x00))],
]
mw = 2.7; mgap = (W - 0.6 - 4*mw) / 3; mx = 0.3; my = 3.25; mh = 1.5; mcx = []
for k, b in enumerate(mol):
    vbox(s, mx, my, mw, mh, b, (AMBER if k == 3 else H_FILL), line=(RGBColor(0xC0,0x7A,0x00) if k == 3 else BLUE), lw=1.25)
    mcx.append(mx + mw); mx += mw + mgap
labels = ["アルドール縮合＋脱水", "逐次縮合", "タンクで伸長"]
for k in range(3):
    arrow(s, mcx[k]+0.03, my+mh/2-0.14, mgap-0.06, 0.28)
    lt = tbx(s, mcx[k]-0.3, my-0.42, mgap+0.6, 0.4)
    P(lt, labels[k], 9.5, color=DBLUE, first=True, align=PP_ALIGN.CENTER)
note = tbx(s, 0.4, 5.1, W-0.8, 1.5)
P(note, "・水が多いEG反応〜濃縮では縮合が進まず、アルデヒドは遊離のまま運ばれる。水が抜ける脱水塔以降で初めて縮合が進み、重い前駆体になってDEG留分へ。", 13, first=True, space=6)
P(note, "・前駆体は製品タンクの長期保管（窒素）で互いに脱水縮合して共役が伸び、共役〜9で450nm＝黄色に着色する。使い切るとAPHA30程度で頭打ち。", 13, space=6)
P(note, "・酸素があると前駆体が有機酸へ酸化され縮合が止まる（窒素タンクで進み、大気接触のドラム・SPで止まる）。温度を下げるとアルデヒドが分解されず下流へ回る。", 13, space=6)

# このスライドをメカニズム(S3)の直後へ移動
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
newid = ids[-1]
sldIdLst.remove(newid)
sldIdLst.insert(3, newid)  # 0-based: 表紙(0)発生事象(1)メカニズム(2)の後＝index3

# ── 巻末：補足資料一覧／用意・確認する資料 ──
def add_table_slide(title, intro, headers, rows, widths):
    s = prs.slides.add_slide(ly)
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx != 0 and ph.placeholder_format.type in (
            PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.CENTER_TITLE):
            ph._element.getparent().remove(ph._element)
    for ph in s.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.text = ""; r = ph.text_frame.paragraphs[0].add_run(); r.text = title
            r.font.size = Pt(24); r.font.bold = True; r.font.color.rgb = BLACK; _ea(r)
    if intro:
        P(tbx(s, 0.5, 1.35, W-1.0, 0.5), intro, 11, color=GRAY, first=True)
    nr = len(rows)+1; nc = len(headers)
    tb = s.shapes.add_table(nr, nc, Emu(int(0.5*IN)), Emu(int(1.9*IN)), Emu(int((W-1.0)*IN)), Emu(int(4.8*IN))).table
    for j, wd in enumerate(widths): tb.columns[j].width = Emu(int(wd*IN))
    def cell(i, j, text, bold=False, fill=WHITE, color=INK, size=10.5):
        c = tb.cell(i, j); c.text = ""; c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = Emu(int(0.05*IN)); c.margin_top = Emu(int(0.02*IN)); c.margin_bottom = Emu(int(0.02*IN))
        p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; _ea(r)
        c.fill.solid(); c.fill.fore_color.rgb = fill
    for j, h in enumerate(headers): cell(0, j, h, bold=True, color=DBLUE, fill=H_FILL, size=11)
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row): cell(i, j, v, fill=(WHITE if i % 2 else RGBColor(0xF7,0xFA,0xFD)), size=10)
    return s

add_table_slide(
    "（巻末）補足資料の一覧　※本編の流れで使うもの・聞かれた時に出すもの",
    "本編：発生事象→推定メカニズム→工程ごとの事実（運転・設備・基礎研）→計画停止対応→S/U時の対応。下記は補足として後ろに置く。",
    ["補足資料（タイトル）", "記載する内容", "状態"],
    [
        ["検査箇所・機器開放要否と結果", "EG精製系の開放箇所と、開けなくてよい理由（脱水塔以降が縮合の場）。C-1503以外に異常なし。", "あり"],
        ["サンプル分析（鉄錆と450nm）", "付着物・タンク異物はほぼ鉄サビ。DEG添加で380nmのみ＝450nmを作らない。", "あり"],
        ["運転対応実績とDEG色相への影響", "NaOH・温度・回収率・還流などの影響（『改善効果』でなく『影響』表現）。温度↓・NaOHはS/U条件の根拠。", "要整理"],
        ["MEG塔リボイラー(E-1501)穴と水分", "2025/10〜の水混入で水分が高めに見えた経緯。修理後は低めで推移。※本筋を混乱させるため聞かれた時のみ。", "あり"],
        ["AALD/FALDバランス", "C-1502で見かけ上ALDが生成（塔間循環）。", "あり"],
        ["UV330nm 活性剤投入有無の比較", "早期色相判断の可否検討。停止前（活性剤なし）と5月（リン酸あり）の比較データ。", "要整理（製造）"],
        ["リン酸の到達位置（C-1503壁面P分析）", "リン酸が効いた場所の絞り込み。XRFでP検出有無。", "結果待ち（基礎研）"],
        ["DEG/TEGリサイクル系のALDバランス", "上流から下流まで関係箇所を網羅したことを示す。", "分析待ち"],
    ],
    [4.2, 6.6, 1.5],
)

add_table_slide(
    "対策会議までに用意・確認する資料（佐野さんと整理）",
    "メカニズム→開放の紐付けを月曜にRD・木村含めて詰める。下記を揃える。",
    ["項目", "内容・狙い", "担当・期限"],
    [
        ["メカニズムの確定", "推定メカニズムに基づき開放箇所を紐づける筋を固める（金属触媒でなくアルデヒド由来）。", "月曜：佐野・高橋・RD・木村"],
        ["リン酸の到達位置", "C-1503壁面のP分析（XRF）。反応場の絞り込み材料。", "基礎研（野田）・近日"],
        ["DEG/TEGリサイクル系バランス", "関係する上流〜下流を網羅したと示すALDバランス。", "S/U会議まで・要レビュー"],
        ["UV330nmの比較データ", "活性剤投入有無での早期判断の可否。比較対象（トラブル前データ）も。", "製造・データ整理"],
        ["初期流動品の監視運用", "一発目はタンクで1週間〜10日UV経時を確認。運転条件変更時の扱いも。", "品証と来週調整"],
        ["運転対応実績の整理", "温度・NaOH・回収率・還流の『影響』。改善効果ありと書かない。", "本編から補足へ"],
    ],
    [3.6, 6.4, 2.3],
)

prs.save(OUT)
import zipfile, collections
names = zipfile.ZipFile(OUT).namelist()
dup = [k for k, v in collections.Counter(names).items() if v > 1]
print("saved:", os.path.normpath(OUT), "| slides:", len(prs.slides._sldIdLst), "| dup:", dup or "なし")
print("S6 filled:", bool(s6), "| S11 filled:", bool(s11))
