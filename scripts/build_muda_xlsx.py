# -*- coding: utf-8 -*-
"""業務効率化／無駄の見える化 Excel生成（段階構成版）。

構成：木村さんのお題 → 定義を開く → 自分ごと（型・基盤がない業務の一覧） → 対策の方向。
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import content_muda as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)

HEADER_FILL = "E6E6E6"; LABEL_FILL = "F2F2F2"; WHITE = "FFFFFF"
INK = "000000"; SUB_INK = "404040"; BORDER_CLR = "808080"
FONT = "IPAGothic"
thin = Side(style="thin", color=BORDER_CLR)
border = Border(left=thin, right=thin, top=thin, bottom=thin)
med = Side(style="medium", color="595959")


def fill(c):
    return PatternFill("solid", fgColor=c)


def est_lines(text, cpl):
    n = 0
    for para in str(text).split("\n"):
        n += max(1, -(-len(para) // cpl))
    return n


wb = Workbook()
ws = wb.active
ws.title = "無駄の見える化"
ws.sheet_view.showGridLines = False

COLW = [("A", 22), ("B", 50), ("C", 46)]
for col, w in COLW:
    ws.column_dimensions[col].width = w
LAST = "C"; NCOL = 3


def borders_row(r):
    for c in range(1, NCOL + 1):
        ws.cell(r, c).border = border


def band(r, text):
    ws.merge_cells(f"A{r}:{LAST}{r}")
    c = ws.cell(r, 1, text)
    c.font = Font(name=FONT, size=11.5, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    borders_row(r)
    ws.row_dimensions[r].height = 22
    return r + 1


def bullets(r, items, cpl=118, ink=INK):
    txt = "・" + "\n・".join(items)
    ws.merge_cells(f"A{r}:{LAST}{r}")
    c = ws.cell(r, 1, txt)
    c.font = Font(name=FONT, size=10, color=ink)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    borders_row(r)
    ws.row_dimensions[r].height = est_lines(txt, cpl) * 15 + 8
    return r + 1


# タイトル
r = 1
ws.merge_cells(f"A{r}:{LAST}{r}")
c = ws.cell(r, 1, C.TITLE)
c.font = Font(name=FONT, size=14, bold=True, color=INK)
c.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[r].height = 26
r += 1
ws.merge_cells(f"A{r}:{LAST}{r}")
c = ws.cell(r, 1, C.SUBTITLE)
c.font = Font(name=FONT, size=9.5, color=SUB_INK)
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
for cidx in range(1, NCOL + 1):
    ws.cell(r, cidx).border = Border(bottom=med)
ws.row_dimensions[r].height = 30
r += 1
c = ws.cell(r, 1, f"（{C.DATE}・平尾）")
c.font = Font(name=FONT, size=9, color=SUB_INK)
r += 1

# お題
r = band(r, "木村さんのお題（6/17 朝会）")
r = bullets(r, C.ODAI)
r += 1
# 定義を開く
r = band(r, "定義をもう少し開くと")
r = bullets(r, C.DEF_OPEN)
r += 1
# 自分ごと A：進め方の無駄（4列＝型/状態/例/対策）
r = band(r, "自分ごとに落とすと ①：進め方の無駄（どう動くか）")
headsA = ["型", "どういう状態が無駄か", "実体験の例", "対策の方向"]
# A表は B列を 状態、C列を 例＋対策にまとめて3列構成で表現（列数を抑える）
heads = ["型", "どういう状態が無駄か", "例 ＋ 対策の方向"]
for i, h in enumerate(heads):
    c = ws.cell(r, 1 + i, h)
    c.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=med)
ws.row_dimensions[r].height = 24
r += 1
cplsA = [20, 50, 46]
for label, state, example, counter in C.SUSUME:
    excol = "例：" + example + "\n→対策：" + counter
    vals = [label, state, excol]
    maxlines = 1
    for i, v in enumerate(vals):
        c = ws.cell(r, 1 + i, v)
        c.font = Font(name=FONT, size=9, bold=(i == 0), color=INK if i != 1 else SUB_INK)
        c.fill = fill(LABEL_FILL if i == 0 else WHITE)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        c.border = border
        maxlines = max(maxlines, est_lines(v, cplsA[i]))
    ws.row_dimensions[r].height = maxlines * 13 + 8
    r += 1
r += 1

# 自分ごと B：仕組み（型・基盤）の無駄（3列＝領域/状態/対策）
r = band(r, "自分ごとに落とすと ②：仕組み（型・基盤）の無駄（毎回ゼロから・その人頼み）")
heads = ["領域", "どういう状態が無駄か（型・基盤がない）", "対策の方向（型・基盤を整備）"]
for i, h in enumerate(heads):
    c = ws.cell(r, 1 + i, h)
    c.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=med)
ws.row_dimensions[r].height = 24
r += 1
cpls = [20, 46, 42]
for area, state, counter in C.SHIKUMI:
    vals = [area, state, counter]
    maxlines = 1
    for i, v in enumerate(vals):
        c = ws.cell(r, 1 + i, v)
        c.font = Font(name=FONT, size=9, bold=(i == 0), color=INK if i != 1 else SUB_INK)
        c.fill = fill(LABEL_FILL if i == 0 else WHITE)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        c.border = border
        maxlines = max(maxlines, est_lines(v, cpls[i]))
    ws.row_dimensions[r].height = maxlines * 13 + 8
    r += 1
r += 1
# 対策の方向
r = band(r, "対策の方向（再現性を上げる）")
r = bullets(r, C.TAISAKU)

# 印刷設定
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.4

wb.properties.creator = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.lastModifiedBy = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.title = "業務効率化／無駄の見える化"

path = os.path.join(OUT, "無駄の見える化_再現性で見る.xlsx")
wb.save(path)
print("saved:", path)
