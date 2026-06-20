# -*- coding: utf-8 -*-
"""解説用Word：製品DEG色相悪化 メカニズム解説（事実→仮説→対応）。配布クリーン。"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import deg_master as M

OUT = os.path.join(os.path.dirname(__file__), "..", "outputs", "製品DEG色相悪化_メカニズム解説.docx")
JP = "游ゴシック"
NAVY = RGBColor(0x1F, 0x38, 0x64)
GRAY = RGBColor(0x55, 0x55, 0x55)

doc = Document()
cp = doc.core_properties
cp.author = ""; cp.last_modified_by = ""; cp.title = "製品DEG色相悪化 メカニズム解説"; cp.comments = ""

# 既定フォント
style = doc.styles["Normal"]
style.font.name = JP; style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn("w:eastAsia"), JP)

def set_font(run, size=10.5, bold=False, color=None):
    run.font.name = JP; run.font.size = Pt(size); run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), JP)
    if color is not None: run.font.color.rgb = color

def title(text):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run(text), 16, True, NAVY)

def sub(text):
    set_font(p_caption(text), 11, True)

def h(text):
    p = doc.add_paragraph(); p.space_before = Pt(6)
    set_font(p.add_run(text), 13, True, NAVY)

def body(text, bullet=False, size=10.5):
    p = doc.add_paragraph(style="List Bullet" if bullet else None)
    set_font(p.add_run(text), size)
    p.paragraph_format.space_after = Pt(3)
    return p

def note(text):
    p = doc.add_paragraph(); set_font(p.add_run(text), 9.5, False, GRAY)
    p.paragraph_format.space_after = Pt(4)

# ── 表紙的ヘッダ ──
title("製品DEG色相悪化　メカニズム解説")
pc = doc.add_paragraph(); pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(pc.add_run("事実の整理 → 推定メカニズム → 対応"), 11, False, GRAY)
doc.add_paragraph()

# 1
h("1. 発生事象（事実）")
for t in M.PHENOMENON: body(t, bullet=True)

# 2
h("2. 着色物質と着色現象")
p = doc.add_paragraph(); set_font(p.add_run("■ 確認した事実"), 11, True, NAVY)
for t in M.COLORANT_FACT: body(t, bullet=True)
p = doc.add_paragraph(); set_font(p.add_run("■ ここから読める推定"), 11, True, GRAY)
for t in M.COLORANT_EST: body(t, bullet=True)

# 3
h("3. " + M.MECH_CAPTION)
note("各段は「運転で見えた事実・基礎研の確認・反応原理」を重ねて導いた最もらしい筋。証明ではないが、観測・分析・原理と整合する。")
for head, txt in M.MECHANISM:
    p = doc.add_paragraph(); set_font(p.add_run(head), 11.5, True, NAVY)
    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(1)
    body(txt)
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
set_font(p.add_run("一筋でいえば："), 11, True)
set_font(p.add_run(M.MECH_ONELINE), 11)
note("※ " + M.MECH_CAVEAT)

# 4
h("4. なぜ今回だけDEGまで到達し着色したか")
note("過去にもアルデヒドが高い時期はあったが今回のような色相悪化は起きていない。アルデヒド単独でなく、反応させてしまう条件が今回そろった。")
for head, txt, act in M.WHY_NOW:
    body("【" + head + "】 " + txt + "　" + act, bullet=True)
note("※ " + M.WATER_NOTE)

# 5
h("5. 設備（鉄サビ・付着物）の関与の評価")
for t in M.RUST: body(t, bullet=True)

# 6
h("6. 開放・検査の整理（仮説駆動）")
body(M.INSPECTION_LOGIC)
for eq, txt in M.INSPECTION:
    body("【" + eq + "】 " + txt, bullet=True)

# 7
h("7. 対応（再発防止・スタートアップ運転条件）")
for head, txt, conf in M.COUNTERMEASURE:
    body("【" + head + "】 " + txt + "（" + conf + "）", bullet=True)

# 8
h("8. 出荷再開に向けた品質管理・早期判定")
for t in M.QC: body(t, bullet=True)

# 9
h("9. 残論点（要確認）")
for head, txt in M.OPEN_ISSUES:
    body("【" + head + "】 " + txt, bullet=True)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", os.path.normpath(OUT))
