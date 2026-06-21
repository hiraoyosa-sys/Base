# -*- coding: utf-8 -*-
"""事実整理Excel：工程別マトリクス中心（工程×①運転 ②RD ③文献→④筋）。薄色・配布クリーン。"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import deg_master as M

OUT = os.path.join(os.path.dirname(__file__), "..", "outputs", "製品DEG色相悪化_事実整理.xlsx")
JP = "游ゴシック"
# 薄色パレット（濃色は使わない。テキストにMCブルーを使う）
BLUE = "005BAB"; DBLUE = "003F7E"; INK = "222222"
H_FILL = "EBEFF2"          # ヘッダ（淡）
PROC = "D7E0E5"            # 工程列・筋（淡青灰）
C_OPE = "E8F1E5"           # ①運転（淡緑）
C_RD = "FBF1DC"            # ②RD（淡黄）
C_LIT = "E9F0F7"           # ③文献（淡青）
C_SUJI = "DCE6EF"          # ④筋（淡青灰）
LINE = "C9D2D8"
thin = Side(style="thin", color=LINE)
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def C(ws, r, c, v, bold=False, size=10.5, color=INK, bg=None, wrap=True, ha="left", va="top"):
    cl = ws.cell(row=r, column=c, value=v)
    cl.font = Font(name=JP, size=size, bold=bold, color=color)
    cl.alignment = Alignment(wrap_text=wrap, vertical=va, horizontal=ha)
    cl.border = border
    if bg: cl.fill = PatternFill("solid", fgColor=bg)
    return cl

def header_row(ws, r, cols, fills=None):
    for i, h in enumerate(cols, 1):
        C(ws, r, i, h, bold=True, size=10.5, color=DBLUE, bg=(fills[i-1] if fills else H_FILL), ha="center", va="center")
    ws.row_dimensions[r].height = 30

wb = openpyxl.Workbook()

# ═══ Sheet1 工程別マトリクス（中心） ═══
ws = wb.active; ws.title = "工程別マトリクス"; ws.sheet_view.showGridLines = False
C(ws, 1, 1, "製品DEG色相悪化　工程別の事実整理　→　1本の筋（仮説）", bold=True, size=13, color=BLUE)
C(ws, 2, 1, "各工程で：①運転で見えた事実　②RDで検証した事実　③文献・この工程で起こる反応　→　④この工程の筋。全工程をトータルすると下段の1本の筋になる。",
  size=9.5, color="555555")
ws.merge_cells("A2:E2")
fills = [PROC, C_OPE, C_RD, C_LIT, C_SUJI]
header_row(ws, 3, M.MATRIX_COLS, fills=fills)
r = 4
for proc, ope, rd, lit, suji in M.PROCESS_MATRIX:
    C(ws, r, 1, proc, bold=True, color=DBLUE, bg=PROC, va="center")
    C(ws, r, 2, ope, bg=C_OPE)
    C(ws, r, 3, rd, bg=C_RD)
    C(ws, r, 4, lit, bg=C_LIT)
    C(ws, r, 5, suji, bold=True, color=DBLUE, bg=C_SUJI)
    ws.row_dimensions[r].height = 92
    r += 1
# 1本の筋
C(ws, r, 1, "トータル＝1本の筋", bold=True, color="FFFFFF", bg=BLUE, va="center", ha="center")
C(ws, r, 2, M.MECH_ONELINE, bold=True, color=DBLUE, bg=C_SUJI)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
ws.row_dimensions[r].height = 60
r += 1
C(ws, r, 1, "注記：本筋は最有力仮説でRDベンチ未再現（タンク長期・低水分・塩基の同時再現が未達）。方向性は観測と整合。", size=9, color="555555")
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
widths = [20, 30, 30, 30, 24]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "B4"

# ═══ Sheet2 発生事象・着色物質 ═══
ws2 = wb.create_sheet("発生事象・着色物質"); ws2.sheet_view.showGridLines = False
C(ws2, 1, 1, "発生事象と着色物質", bold=True, size=13, color=BLUE)
header_row(ws2, 2, ["区分", "内容"])
rows = [("発生事象（事実）", t) for t in M.PHENOMENON] \
     + [("着色物質・現象（事実）", t) for t in M.COLORANT_FACT] \
     + [("着色物質（推定）", t) for t in M.COLORANT_EST]
r = 3; prev = None
for kind, t in rows:
    C(ws2, r, 1, kind if kind != prev else "", bold=True, color=DBLUE, bg=PROC, va="center"); prev = kind
    C(ws2, r, 2, t)
    ws2.row_dimensions[r].height = 40; r += 1
ws2.column_dimensions["A"].width = 20; ws2.column_dimensions["B"].width = 100
ws2.freeze_panes = "A3"

# ═══ Sheet3 なぜ今回 ═══
ws3 = wb.create_sheet("なぜ今回起きたか"); ws3.sheet_view.showGridLines = False
C(ws3, 1, 1, "なぜ今回だけDEGまで到達し着色したか（条件の重なり）", bold=True, size=13, color=BLUE)
header_row(ws3, 2, ["要因（位置づけ）", "内容", "対応"])
r = 3
for head, txt, act in M.WHY_NOW:
    C(ws3, r, 1, head, bold=True, color=DBLUE, bg=PROC, va="center"); C(ws3, r, 2, txt); C(ws3, r, 3, act, color=BLUE, bold=True)
    ws3.row_dimensions[r].height = 44; r += 1
C(ws3, r, 1, "水分の注記", bold=True, color=DBLUE, bg=C_RD, va="center")
C(ws3, r, 2, M.WATER_NOTE, bg=C_RD); C(ws3, r, 3, "", bg=C_RD); ws3.row_dimensions[r].height = 56
ws3.column_dimensions["A"].width = 30; ws3.column_dimensions["B"].width = 74; ws3.column_dimensions["C"].width = 24
ws3.freeze_panes = "A3"

# ═══ Sheet4 設備・開放 ═══
ws4 = wb.create_sheet("設備・開放"); ws4.sheet_view.showGridLines = False
C(ws4, 1, 1, "設備（鉄サビ・付着物）の評価と開放・検査", bold=True, size=13, color=BLUE)
C(ws4, 2, 1, "■ 鉄サビ・付着物の評価", bold=True, size=11, color=BLUE)
ws4.merge_cells("A2:B2")
r = 3
for t in M.RUST:
    C(ws4, r, 1, t); ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); ws4.row_dimensions[r].height = 36; r += 1
C(ws4, r, 1, "■ 開放・検査（機器ごと）", bold=True, size=11, color=BLUE); ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); r += 1
header_row(ws4, r, ["機器", "開放理由・結果"]); r += 1
for eq, txt in M.INSPECTION:
    C(ws4, r, 1, eq, bold=True, color=DBLUE, bg=PROC, va="center"); C(ws4, r, 2, txt); ws4.row_dimensions[r].height = 32; r += 1
C(ws4, r, 1, M.INSPECTION_LOGIC, color="555555", size=9.5); ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); ws4.row_dimensions[r].height = 40
ws4.column_dimensions["A"].width = 32; ws4.column_dimensions["B"].width = 86

# ═══ Sheet5 対応 ═══
ws5 = wb.create_sheet("対応"); ws5.sheet_view.showGridLines = False
C(ws5, 1, 1, "対応（再発防止・スタートアップ運転条件）", bold=True, size=13, color=BLUE)
header_row(ws5, 2, ["項目", "内容", "状況"])
r = 3
for head, txt, conf in M.COUNTERMEASURE:
    C(ws5, r, 1, head, bold=True, color=DBLUE, bg=PROC, va="center"); C(ws5, r, 2, txt); C(ws5, r, 3, conf, va="center", ha="center")
    ws5.row_dimensions[r].height = 42; r += 1
ws5.column_dimensions["A"].width = 24; ws5.column_dimensions["B"].width = 82; ws5.column_dimensions["C"].width = 14
ws5.freeze_panes = "A3"

# ═══ Sheet6 品質管理・残論点 ═══
ws6 = wb.create_sheet("品質管理・残論点"); ws6.sheet_view.showGridLines = False
C(ws6, 1, 1, "出荷再開に向けた品質管理・早期判定と残論点", bold=True, size=13, color=BLUE)
C(ws6, 2, 1, "■ 品質管理・早期判定", bold=True, size=11, color=BLUE); ws6.merge_cells("A2:B2")
r = 3
for t in M.QC:
    C(ws6, r, 1, t); ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); ws6.row_dimensions[r].height = 34; r += 1
C(ws6, r, 1, "■ 残論点（要確認）", bold=True, size=11, color=BLUE); ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); r += 1
header_row(ws6, r, ["論点", "確認の方向"]); r += 1
for head, txt in M.OPEN_ISSUES:
    C(ws6, r, 1, head, bold=True, color=DBLUE, bg=PROC, va="center"); C(ws6, r, 2, txt); ws6.row_dimensions[r].height = 38; r += 1
ws6.column_dimensions["A"].width = 24; ws6.column_dimensions["B"].width = 86

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("saved:", os.path.normpath(OUT), "| sheets:", wb.sheetnames)
