# -*- coding: utf-8 -*-
"""DEG着色シナリオ 横フロー整理（最新版）Excel生成。

横＝プロセス工程、縦＝トピック行。各セルは「その塔で何を言うか」を列内に割り付ける。
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）。
別シートに 確認事実(F)・6/16更新点・要検証。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import content as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)

# ---- モノクロ配色 ----
HEADER_FILL = "E6E6E6"   # 工程ヘッダ（薄グレー）
LABEL_FILL  = "F2F2F2"   # 行ラベル列（より薄いグレー）
SUB_FILL    = "F7F7F7"   # 補助行の薄い陰
WHITE       = "FFFFFF"
INK         = "000000"
SUB_INK     = "404040"
BORDER_CLR  = "808080"

FONT = "IPAGothic"
thin = Side(style="thin", color=BORDER_CLR)
border = Border(left=thin, right=thin, top=thin, bottom=thin)
med = Side(style="medium", color="595959")


def fill(c):
    return PatternFill("solid", fgColor=c)


def est_lines(text, cpl):
    n = 0
    for para in text.split("\n"):
        n += max(1, -(-len(para) // cpl))
    return n


wb = Workbook()
ws = wb.active
ws.title = "横フロー整理"
ws.sheet_view.showGridLines = False

NSTEP = len(C.STEPS)
LBL_W = 22
STEP_W = 33
ws.column_dimensions["A"].width = LBL_W
for i in range(NSTEP):
    ws.column_dimensions[get_column_letter(2 + i)].width = STEP_W
last_col = 1 + NSTEP
last_letter = get_column_letter(last_col)
STEP_CPL = int(STEP_W / 2.05)
WIDE_CPL = int((LBL_W + STEP_W * NSTEP) / 2.05)

r = 1
# タイトル（塗りつぶしなし・黒太字＋下罫線）
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, C.TITLE)
c.font = Font(name=FONT, size=15, bold=True, color=INK)
c.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[r].height = 26
r += 1
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, f"（{C.DATE} 時点）")
c.font = Font(name=FONT, size=9, color=SUB_INK)
c.alignment = Alignment(horizontal="right", vertical="center")
for cidx in range(1, last_col + 1):
    ws.cell(r, cidx).border = Border(bottom=med)
ws.row_dimensions[r].height = 14
r += 1

# 工程ヘッダ
HEADER_ROW = r
hc = ws.cell(r, 1, "工程 →")
hc.font = Font(name=FONT, size=10, bold=True, color=INK)
hc.fill = fill(HEADER_FILL)
hc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
hc.border = border
for i, (name, hl) in enumerate(C.STEPS):
    cell = ws.cell(r, 2 + i, name)
    cell.font = Font(name=FONT, size=10, bold=True, color=INK)
    cell.fill = fill(HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=med)
ws.row_dimensions[r].height = 34
r += 1

# トピック行
SHADED = {"rust", "drug", "facts", "todo"}  # 補助情報行はごく薄い陰
for label, key in C.ROW_DEFS:
    shade = SUB_FILL if key in SHADED else WHITE
    lab = ws.cell(r, 1, label)
    lab.font = Font(name=FONT, size=9, bold=True, color=INK)
    lab.fill = fill(LABEL_FILL)
    lab.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    lab.border = border
    maxlines = 1
    for i in range(NSTEP):
        txt = C.COLS[i][key]
        cell = ws.cell(r, 2 + i, txt)
        cell.font = Font(name=FONT, size=8.5, color=INK if key in ("kept", "kept_reason") else SUB_INK)
        cell.fill = fill(shade)
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        cell.border = border
        maxlines = max(maxlines, est_lines(txt, STEP_CPL))
    ws.row_dimensions[r].height = maxlines * 12.5 + 7
    r += 1

# 一本化した筋（全幅・総括）
r += 1
lab = ws.cell(r, 1, "一本化した筋")
lab.font = Font(name=FONT, size=10, bold=True, color=INK)
lab.fill = fill(HEADER_FILL)
lab.alignment = Alignment(horizontal="left", vertical="center", indent=1)
lab.border = border
ws.merge_cells(f"B{r}:{last_letter}{r}")
c = ws.cell(r, 2, C.SUJI)
c.font = Font(name=FONT, size=9.5, color=INK)
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
for cidx in range(1, last_col + 1):
    ws.cell(r, cidx).border = border
ws.row_dimensions[r].height = est_lines(C.SUJI, WIDE_CPL - 8) * 15 + 10

# 印刷設定
ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.4
ws.freeze_panes = "B" + str(HEADER_ROW + 1)


# ====== シート2：事実・更新・要検証 ======
ws2 = wb.create_sheet("事実と更新・要検証")
ws2.sheet_view.showGridLines = False
ws2.column_dimensions["A"].width = 14
ws2.column_dimensions["B"].width = 104


def s2_band(r, text):
    ws2.merge_cells(f"A{r}:B{r}")
    c = ws2.cell(r, 1, text)
    c.font = Font(name=FONT, size=12, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border = border
    ws2.cell(r, 2).border = border
    ws2.row_dimensions[r].height = 22
    return r + 1


def s2_kv(r, k, v, kbold=False):
    kc = ws2.cell(r, 1, k)
    kc.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    kc.fill = fill(LABEL_FILL)
    kc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    kc.border = border
    vc = ws2.cell(r, 2, v)
    vc.font = Font(name=FONT, size=9.5, color=INK, bold=kbold)
    vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    vc.border = border
    ws2.row_dimensions[r].height = max(26, est_lines(v, 52) * 15 + 7)
    return r + 1


r = 1
r = s2_band(r, "6/16での主な更新（6/15引き継ぎからの差分）")
for k, v in C.UPDATES:
    r = s2_kv(r, k, v)
r += 1
r = s2_band(r, "確認された事実（F番号）")
for k, v in C.FACTS:
    r = s2_kv(r, k, v, kbold=(k == "F2"))
r += 1
r = s2_band(r, "要検証・齟齬（要確認）")
for i, v in enumerate(C.OPEN, 1):
    r = s2_kv(r, str(i), v)

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
