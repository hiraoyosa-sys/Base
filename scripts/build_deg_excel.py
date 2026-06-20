# -*- coding: utf-8 -*-
"""事実整理Excel：製品DEG色相悪化。配布クリーン（外部引用・AI痕跡なし・標準体裁）。"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import deg_master as M

OUT = os.path.join(os.path.dirname(__file__), "..", "outputs", "製品DEG色相悪化_事実整理.xlsx")
JP = "游ゴシック"
HEAD = "1F3864"; SUB = "D9E1F2"; LINE = "BFBFBF"
thin = Side(style="thin", color=LINE)
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def C(ws, r, c, v, bold=False, size=10.5, color="222222", bg=None, wrap=True, ha="left", va="top"):
    cl = ws.cell(row=r, column=c, value=v)
    cl.font = Font(name=JP, size=size, bold=bold, color=color)
    cl.alignment = Alignment(wrap_text=wrap, vertical=va, horizontal=ha)
    cl.border = border
    if bg: cl.fill = PatternFill("solid", fgColor=bg)
    return cl

def header_row(ws, r, cols):
    for i, h in enumerate(cols, 1):
        C(ws, r, i, h, bold=True, size=11, color="FFFFFF", bg=HEAD, ha="center", va="center")
    ws.row_dimensions[r].height = 26

wb = openpyxl.Workbook()

# ── Sheet1 事実の整理 ──
ws = wb.active; ws.title = "事実の整理"; ws.sheet_view.showGridLines = False
C(ws, 1, 1, "製品DEG色相悪化　事実の整理", bold=True, size=13, color=HEAD)
header_row(ws, 2, ["区分", "内容"])
rows = []
for t in M.PHENOMENON: rows.append(("発生事象（事実）", t))
for t in M.COLORANT_FACT: rows.append(("着色物質・現象（事実）", t))
for t in M.COLORANT_EST: rows.append(("着色物質（推定）", t))
r = 3
prev = None
for kind, t in rows:
    C(ws, r, 1, kind if kind != prev else "", bold=True, bg=SUB, va="center"); prev = kind
    C(ws, r, 2, t)
    ws.row_dimensions[r].height = 42; r += 1
ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 96
ws.freeze_panes = "A3"

# ── Sheet2 推定メカニズム ──
ws2 = wb.create_sheet("推定メカニズム"); ws2.sheet_view.showGridLines = False
C(ws2, 1, 1, M.MECH_CAPTION, bold=True, size=13, color=HEAD)
header_row(ws2, 2, ["段", "内容"])
r = 3
for head, txt in M.MECHANISM:
    C(ws2, r, 1, head, bold=True, bg=SUB, va="center")
    C(ws2, r, 2, txt)
    ws2.row_dimensions[r].height = 64; r += 1
C(ws2, r, 1, "一筋", bold=True, bg="EAEFF7", va="center")
C(ws2, r, 2, M.MECH_ONELINE, bold=False, bg="EAEFF7")
ws2.row_dimensions[r].height = 64; r += 1
C(ws2, r, 1, "注記", bold=True, bg="FFF2CC", va="center")
C(ws2, r, 2, M.MECH_CAVEAT, bold=False, bg="FFF2CC")
ws2.row_dimensions[r].height = 48
ws2.column_dimensions["A"].width = 34; ws2.column_dimensions["B"].width = 90
ws2.freeze_panes = "A3"

# ── Sheet3 なぜ今回 ──
ws3 = wb.create_sheet("なぜ今回起きたか"); ws3.sheet_view.showGridLines = False
C(ws3, 1, 1, "なぜ今回だけDEGまで到達し着色したか（条件の重なり）", bold=True, size=13, color=HEAD)
header_row(ws3, 2, ["要因（位置づけ）", "内容", "対応"])
r = 3
for head, txt, act in M.WHY_NOW:
    C(ws3, r, 1, head, bold=True, bg=SUB, va="center"); C(ws3, r, 2, txt); C(ws3, r, 3, act, color="1F3864", bold=True)
    ws3.row_dimensions[r].height = 44; r += 1
C(ws3, r, 1, "水分の注記", bold=True, bg="FFF2CC", va="center")
C(ws3, r, 2, M.WATER_NOTE, bg="FFF2CC"); C(ws3, r, 3, "", bg="FFF2CC"); ws3.row_dimensions[r].height = 56
ws3.column_dimensions["A"].width = 30; ws3.column_dimensions["B"].width = 74; ws3.column_dimensions["C"].width = 24
ws3.freeze_panes = "A3"

# ── Sheet4 設備の関与・開放 ──
ws4 = wb.create_sheet("設備の関与・開放"); ws4.sheet_view.showGridLines = False
C(ws4, 1, 1, "設備（鉄サビ・付着物）の関与の評価と開放・検査の整理", bold=True, size=13, color=HEAD)
C(ws4, 2, 1, "■ 鉄サビ・付着物の評価", bold=True, size=11, color=HEAD)
r = 3
for t in M.RUST:
    C(ws4, r, 1, t); ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws4.row_dimensions[r].height = 38; r += 1
C(ws4, r, 1, "■ 開放・検査（機器ごと）", bold=True, size=11, color=HEAD)
ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); r += 1
header_row(ws4, r, ["機器", "開放理由・結果"]); r += 1
for eq, txt in M.INSPECTION:
    C(ws4, r, 1, eq, bold=True, bg=SUB, va="center"); C(ws4, r, 2, txt)
    ws4.row_dimensions[r].height = 34; r += 1
C(ws4, r, 1, M.INSPECTION_LOGIC, color="555555", size=9.5)
ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); ws4.row_dimensions[r].height = 40
ws4.column_dimensions["A"].width = 34; ws4.column_dimensions["B"].width = 80

# ── Sheet5 対応 ──
ws5 = wb.create_sheet("対応"); ws5.sheet_view.showGridLines = False
C(ws5, 1, 1, "対応（再発防止・スタートアップ運転条件）", bold=True, size=13, color=HEAD)
header_row(ws5, 2, ["項目", "内容", "位置づけ"])
r = 3
for head, txt, conf in M.COUNTERMEASURE:
    C(ws5, r, 1, head, bold=True, bg=SUB, va="center"); C(ws5, r, 2, txt); C(ws5, r, 3, conf, va="center")
    ws5.row_dimensions[r].height = 42; r += 1
ws5.column_dimensions["A"].width = 26; ws5.column_dimensions["B"].width = 76; ws5.column_dimensions["C"].width = 16
ws5.freeze_panes = "A3"

# ── Sheet6 品質管理・残論点 ──
ws6 = wb.create_sheet("品質管理・残論点"); ws6.sheet_view.showGridLines = False
C(ws6, 1, 1, "出荷再開に向けた品質管理・早期判定と残論点", bold=True, size=13, color=HEAD)
C(ws6, 2, 1, "■ 品質管理・早期判定", bold=True, size=11, color=HEAD)
ws6.merge_cells(start_row=2, start_column=1, end_row=2, end_column=2)
r = 3
for t in M.QC:
    C(ws6, r, 1, t); ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws6.row_dimensions[r].height = 34; r += 1
C(ws6, r, 1, "■ 残論点（要確認）", bold=True, size=11, color=HEAD)
ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); r += 1
header_row(ws6, r, ["論点", "確認の方向"]); r += 1
for head, txt in M.OPEN_ISSUES:
    C(ws6, r, 1, head, bold=True, bg=SUB, va="center"); C(ws6, r, 2, txt)
    ws6.row_dimensions[r].height = 38; r += 1
ws6.column_dimensions["A"].width = 26; ws6.column_dimensions["B"].width = 80

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("saved:", os.path.normpath(OUT), "| sheets:", wb.sheetnames)
