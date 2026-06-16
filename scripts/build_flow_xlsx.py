# -*- coding: utf-8 -*-
"""DEG着色シナリオ 横フロー整理（詳細版）Excel生成。

横＝プロセスフロー、縦＝4段。下部に
 リン酸の効き場所／一本化した筋／機構の確定事項／鉄の2つの顔 を文章で付す。
別シートに 確認事実(F1-F10)＋塔ごとの場所事実。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import content as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)

# ---- 配色 ----
C_TITLE   = "1F3864"
C_HEAD    = "2E5496"
C_HEAD_HL = "9E3A26"
C_KEPT    = "DDEBF7"
C_REASON  = "E2EFDA"
C_DROP    = "F2F2F2"
C_DROPWHY = "D9D9D9"
C_SECT    = "44546A"
C_IRON_OK = "E2EFDA"
C_IRON_NG = "FCE4D6"
WHITE = "FFFFFF"

FONT = "IPAGothic"
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)


def fill(c):
    return PatternFill("solid", fgColor=c)


def est_lines(text, chars_per_line):
    """全角想定でセル内行数をざっくり見積もる（改行コードも考慮）。"""
    n = 0
    for para in text.split("\n"):
        n += max(1, -(-len(para) // chars_per_line))
    return n


wb = Workbook()
ws = wb.active
ws.title = "フロー整理"
ws.sheet_view.showGridLines = False

NSTEP = len(C.STEPS)
LBL_W = 24
STEP_W = 32
ws.column_dimensions["A"].width = LBL_W
for i in range(NSTEP):
    ws.column_dimensions[get_column_letter(2 + i)].width = STEP_W
last_col = 1 + NSTEP
last_letter = get_column_letter(last_col)
# 折返し見積り用の1行あたり全角文字数
STEP_CPL = int(STEP_W / 2.05)   # 工程セル
WIDE_CPL = int((LBL_W + STEP_W * NSTEP) / 2.05)  # 全幅マージセル
PT = 14.5  # 1行あたり高さ(pt)


def band(ws, r, text, fillc=C_SECT, size=11, color=WHITE, h=22):
    ws.merge_cells(f"A{r}:{last_letter}{r}")
    c = ws.cell(r, 1, text)
    c.font = Font(name=FONT, size=size, bold=True, color=color)
    c.fill = fill(fillc)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = h
    return r + 1


def kv_row(ws, r, k, v, kfill=C_DROP, vfill=WHITE, klen=None, kcolor="333333"):
    """A列=ラベル, B:last=本文 の行。高さは本文行数で自動。"""
    kc = ws.cell(r, 1, k)
    kc.font = Font(name=FONT, size=9.5, bold=True, color=kcolor)
    kc.fill = fill(kfill)
    kc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    kc.border = border
    ws.merge_cells(f"B{r}:{last_letter}{r}")
    vc = ws.cell(r, 2, v)
    vc.font = Font(name=FONT, size=9.5, color="222222")
    vc.fill = fill(vfill)
    vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for cidx in range(1, last_col + 1):
        ws.cell(r, cidx).border = border
    cpl = WIDE_CPL - (klen or 8)
    lines = est_lines(v, max(20, cpl))
    ws.row_dimensions[r].height = max(28, lines * PT + 8)
    return r + 1


r = 1
# タイトル
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, C.TITLE)
c.font = Font(name=FONT, size=16, bold=True, color=WHITE)
c.fill = fill(C_TITLE)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[r].height = 30
r += 1
# 日付
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, C.DATE)
c.font = Font(name=FONT, size=10, color="595959")
c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
ws.row_dimensions[r].height = 16
r += 1

# 工程ヘッダ
HEADER_ROW = r
lab = ws.cell(r, 1, "工程フロー →")
lab.font = Font(name=FONT, size=10, bold=True, color=WHITE)
lab.fill = fill(C_SECT)
lab.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
lab.border = border
for i, (name, hl) in enumerate(C.STEPS):
    cell = ws.cell(r, 2 + i, name)
    cell.font = Font(name=FONT, size=10.5, bold=True, color=WHITE)
    cell.fill = fill(C_HEAD_HL if hl else C_HEAD)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border
ws.row_dimensions[r].height = 40
r += 1

# 4段（行高は各段の最長セルから自動）
ROW_FILLS = [C_KEPT, C_REASON, C_DROP, C_DROPWHY]
for t in range(4):
    lab = ws.cell(r, 1, C.ROW_LABELS[t])
    lab.font = Font(name=FONT, size=9.5, bold=True, color="333333")
    lab.fill = fill(ROW_FILLS[t])
    lab.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    lab.border = border
    maxlines = 1
    for i in range(NSTEP):
        text = C.DATA[i][t]
        cell = ws.cell(r, 2 + i, text)
        cell.font = Font(name=FONT, size=9, color="222222")
        cell.fill = fill(ROW_FILLS[t])
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        cell.border = border
        maxlines = max(maxlines, est_lines(text, STEP_CPL))
    ws.row_dimensions[r].height = maxlines * 13.5 + 8
    r += 1

r += 1
# リン酸の効き場所
r = band(ws, r, "リン酸の効き場所（絞り込み済み・確定未）")
for k, v in C.PHOS:
    r = kv_row(ws, r, k, v, kfill=C_DROP, klen=7)

r += 1
# 一本化した筋
r = band(ws, r, "一本化した筋")
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, C.SUJI)
c.font = Font(name=FONT, size=10, color="222222")
c.fill = fill(C_REASON)
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
for cidx in range(1, last_col + 1):
    ws.cell(r, cidx).border = border
ws.row_dimensions[r].height = est_lines(C.SUJI, WIDE_CPL) * PT + 10
r += 1

r += 1
# 機構の確定事項
r = band(ws, r, "機構の確定事項（理由欄の裏付け）")
for k, v in C.MECH:
    r = kv_row(ws, r, k, v, kfill=C_KEPT, klen=12)

r += 1
# 鉄の2つの顔
r = band(ws, r, "鉄の役割＝2つの顔（今回の深掘り）")
IRON_FILLS = {0: C_IRON_OK, 1: C_IRON_NG}
for idx, (k, v) in enumerate(C.IRON):
    vfill = IRON_FILLS.get(idx, WHITE)
    kfill = IRON_FILLS.get(idx, C_DROP)
    r = kv_row(ws, r, k, v, kfill=kfill, vfill=vfill, klen=12)

# 印刷設定
ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.4

# ====== シート2：塔ごとの場所事実 ======
ws2 = wb.create_sheet("塔ごとの場所事実")
ws2.sheet_view.showGridLines = False
ws2.column_dimensions["A"].width = 18
ws2.column_dimensions["B"].width = 95
r = 1
ws2.merge_cells(f"A{r}:B{r}")
c = ws2.cell(r, 1, "確認された事実（F番号）")
c.font = Font(name=FONT, size=14, bold=True, color=WHITE)
c.fill = fill(C_TITLE)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[r].height = 28
r += 1
for k, v in C.FACTS:
    kc = ws2.cell(r, 1, k)
    kc.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    kc.fill = fill(C_HEAD)
    kc.alignment = Alignment(horizontal="center", vertical="center")
    kc.border = border
    hl = (k == "F2")
    vc = ws2.cell(r, 2, v)
    vc.font = Font(name=FONT, size=10, color="222222", bold=hl)
    vc.fill = fill(C_REASON if hl else WHITE)
    vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    vc.border = border
    ws2.row_dimensions[r].height = max(28, est_lines(v, 46) * 15 + 8)
    r += 1

r += 1
ws2.merge_cells(f"A{r}:B{r}")
c = ws2.cell(r, 1, "塔ごとの場所事実（追記枠）")
c.font = Font(name=FONT, size=12, bold=True, color=WHITE)
c.fill = fill(C_SECT)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[r].height = 24
r += 1
for k, v in C.PLACE:
    kc = ws2.cell(r, 1, k)
    kc.font = Font(name=FONT, size=9.5, bold=True, color="333333")
    kc.fill = fill(C_DROP)
    kc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    kc.border = border
    vc = ws2.cell(r, 2, v)
    vc.font = Font(name=FONT, size=9.5, color="222222")
    vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    vc.border = border
    ws2.row_dimensions[r].height = max(28, est_lines(v, 46) * 15 + 8)
    r += 1

ws2.page_setup.orientation = "landscape"
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 0
ws2.sheet_properties.pageSetUpPr.fitToPage = True

# ---- メタデータ（作成者）----
wb.properties.creator = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.lastModifiedBy = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.title = "DEG着色シナリオ 横フロー整理"
wb.properties.description = ""
wb.properties.keywords = ""

path = os.path.join(OUT, "DEG着色_フロー整理.xlsx")
wb.save(path)
print("saved:", path)
