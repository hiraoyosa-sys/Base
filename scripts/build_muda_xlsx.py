# -*- coding: utf-8 -*-
"""業務効率化／無駄の見える化 Excel生成（無駄の型リスト版）。

シート：表紙・考え方／無駄の型一覧（型・状態・実体験の例・対策）／共通点・発信メモ。
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）で content 系資料と統一。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import content_muda as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)

HEADER_FILL = "E6E6E6"; LABEL_FILL = "F2F2F2"; SUB_FILL = "F7F7F7"
WHITE = "FFFFFF"; INK = "000000"; SUB_INK = "404040"; BORDER_CLR = "808080"
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
ws.title = "無駄の型"
ws.sheet_view.showGridLines = False

# 列：型 / カテゴリ / どういう状態が無駄か / 実体験の例 / 対策の方向
COLW = [("A", 22), ("B", 13), ("C", 34), ("D", 46), ("E", 38)]
for col, w in COLW:
    ws.column_dimensions[col].width = w
LAST = "E"; NCOL = 5


def borders_row(r):
    for c in range(1, NCOL + 1):
        ws.cell(r, c).border = border


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
c = ws.cell(r, 1, f"（{C.DATE} 時点・木村さん 6/17発信への対応／自分の振り返り）")
c.font = Font(name=FONT, size=9, color=SUB_INK)
r += 1

# 冒頭の考え
ws.merge_cells(f"A{r}:{LAST}{r}")
c = ws.cell(r, 1, "無駄とは何か（私の考え）")
c.font = Font(name=FONT, size=11.5, bold=True, color=INK)
c.fill = fill(HEADER_FILL)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
borders_row(r)
ws.row_dimensions[r].height = 22
r += 1
intro_text = "・" + "\n・".join(C.INTRO)
ws.merge_cells(f"A{r}:{LAST}{r}")
c = ws.cell(r, 1, intro_text)
c.font = Font(name=FONT, size=10, color=INK)
c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
borders_row(r)
ws.row_dimensions[r].height = est_lines(intro_text, 150) * 15 + 8
r += 1

# 型一覧ヘッダ
heads = ["無駄の型", "カテゴリ", "どういう状態が無駄か", "私の実体験の例（複数PJ・過去〜直近）", "対策の方向（行動・順序）"]
for i, h in enumerate(heads):
    c = ws.cell(r, 1 + i, h)
    c.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=med)
ws.row_dimensions[r].height = 26
r += 1

cpls = [20, 11, 30, 42, 34]
for label, cat, state, examples, counter in C.MUDA_TYPES:
    ex_text = "・" + "\n・".join(examples)
    vals = [label, cat, state, ex_text, counter]
    maxlines = 1
    for i, v in enumerate(vals):
        c = ws.cell(r, 1 + i, v)
        bold = (i == 0)
        c.font = Font(name=FONT, size=9, bold=bold, color=INK if i in (0, 2, 4) else SUB_INK)
        c.fill = fill(LABEL_FILL if i == 0 else WHITE)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        c.border = border
        maxlines = max(maxlines, est_lines(v, cpls[i]))
    ws.row_dimensions[r].height = maxlines * 13 + 8
    r += 1

# 共通点
r += 1
ws.merge_cells(f"A{r}:{LAST}{r}")
c = ws.cell(r, 1, "強いて共通点を言えば（※1個直せば終わり、ではない）")
c.font = Font(name=FONT, size=11, bold=True, color=INK)
c.fill = fill(HEADER_FILL)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
borders_row(r)
ws.row_dimensions[r].height = 22
r += 1
ws.merge_cells(f"A{r}:{LAST}{r}")
c = ws.cell(r, 1, C.COMMON)
c.font = Font(name=FONT, size=10, color=INK)
c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
borders_row(r)
ws.row_dimensions[r].height = est_lines(C.COMMON, 150) * 15 + 8
r += 1

# 発信メモ
r += 1
for k, v in C.RELEASE:
    kc = ws.cell(r, 1, k)
    kc.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    kc.fill = fill(LABEL_FILL)
    kc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    kc.border = border
    ws.merge_cells(f"B{r}:{LAST}{r}")
    vc = ws.cell(r, 2, v)
    vc.font = Font(name=FONT, size=9.5, color=SUB_INK)
    vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    borders_row(r)
    ws.row_dimensions[r].height = est_lines(v, 120) * 15 + 6
    r += 1

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
