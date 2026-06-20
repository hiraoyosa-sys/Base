# -*- coding: utf-8 -*-
"""EG品質／DEG色相 事実整理 Excel生成。

ルール準拠：マス結合しない／標準色（白地・黒文字・薄グレー見出し）のみ／
セルに数式を入れない（先頭 = や @ を出さない）／勝手な語（採用 等）を付けない／装飾を盛らない。
作成者メタは平尾名義。AI痕跡・外部文献は入れない。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import eg_content as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)

FONT = "Meiryo"
INK = "000000"
SUB = "404040"
HEAD_FILL = "D9D9D9"   # 見出し（薄グレー）
TITLE_FILL = "BFBFBF"  # セクション帯（やや濃いグレー）
ALT_FILL = "F2F2F2"    # 交互行
WHITE = "FFFFFF"
BORDER = "808080"

thin = Side(style="thin", color=BORDER)
box = Border(left=thin, right=thin, top=thin, bottom=thin)


def fillc(c):
    return PatternFill("solid", fgColor=c)


def estlines(text, cpl):
    n = 0
    for para in str(text).split("\n"):
        n += max(1, -(-len(para) // cpl))
    return n


def title(ws, r, text, ncol):
    """セクション帯（結合せず、1行ぶんを薄グレーで塗る）。"""
    for c in range(1, ncol + 1):
        cell = ws.cell(r, c)
        cell.fill = fillc(TITLE_FILL)
        cell.border = box
    cell = ws.cell(r, 1, text)
    cell.font = Font(name=FONT, size=11, bold=True, color=INK)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    return r + 1


def header_row(ws, r, headers, widths=None):
    for i, h in enumerate(headers, 1):
        cell = ws.cell(r, i, h)
        cell.font = Font(name=FONT, size=9.5, bold=True, color=INK)
        cell.fill = fillc(HEAD_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = box
    ws.row_dimensions[r].height = 26
    return r + 1


def data_row(ws, r, values, cpls, aligns=None, alt=False, strong_first=False):
    aligns = aligns or ["left"] * len(values)
    maxlines = 1
    for i, v in enumerate(values, 1):
        cell = ws.cell(r, i, v)
        bold = strong_first and i == 1
        cell.font = Font(name=FONT, size=9, color=INK, bold=bold)
        cell.fill = fillc(ALT_FILL if alt else WHITE)
        cell.alignment = Alignment(horizontal=aligns[i - 1], vertical="top", wrap_text=True, indent=1)
        cell.border = box
        maxlines = max(maxlines, estlines(v, cpls[i - 1]))
    ws.row_dimensions[r].height = maxlines * 13.5 + 6
    return r + 1


def setw(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def page(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4


wb = Workbook()

# ============ Sheet 0: 表紙・目的 ============
ws = wb.active
ws.title = "0.目的"
setw(ws, [4, 110])
page(ws)
c = ws.cell(1, 1, C.TITLE)
c.font = Font(name=FONT, size=16, bold=True, color=INK)
c.alignment = Alignment(vertical="center")
ws.merge_cells  # noqa: B018  (使わない方針の明示。結合はしない)
ws.cell(2, 1)
c = ws.cell(2, 2, C.SUBTITLE)
c.font = Font(name=FONT, size=10, color=SUB)
c = ws.cell(3, 2, f"{C.DATE} 時点")
c.font = Font(name=FONT, size=9, color=SUB)
r = 5
r = title(ws, r, "この資料の目的", 2)
for line in C.PURPOSE:
    cell = ws.cell(r, 2, line)
    cell.font = Font(name=FONT, size=10, color=INK)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    cell.border = box
    ws.cell(r, 1).border = box
    ws.row_dimensions[r].height = max(20, estlines(line, 80) * 15 + 5)
    r += 1
r += 1
r = title(ws, r, "結論（先に要点）", 2)
for line in C.SUMMARY:
    cell = ws.cell(r, 2, "・" + line)
    cell.font = Font(name=FONT, size=10, color=INK)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    cell.border = box
    ws.cell(r, 1).border = box
    ws.row_dimensions[r].height = max(20, estlines(line, 78) * 15 + 5)
    r += 1

# ============ Sheet 1: 現状 ============
ws = wb.create_sheet("1.現状")
setw(ws, [16, 52, 52])
page(ws)
r = 1
r = title(ws, r, "製品別の品質状況（事実）", 3)
r = header_row(ws, r, ["製品／系", "状況", "特徴的な事象"])
for i, (a, b, c2) in enumerate(C.PRODUCT_STATUS):
    r = data_row(ws, r, [a, b, c2], [16, 30, 30], alt=(i % 2 == 1), strong_first=True)
r += 1
r = title(ws, r, "顧客の対応窓口", 3)
r = header_row(ws, r, ["顧客", "対応の対象", ""])
for i, (a, b) in enumerate(C.CUSTOMER_MAP):
    r = data_row(ws, r, [a, b, ""], [16, 60, 4], alt=(i % 2 == 1), strong_first=True)

# ============ Sheet 2: 運転対応実績 ============
ws = wb.create_sheet("2.運転対応実績")
setw(ws, [40, 18, 16, 16])
page(ws)
r = 1
r = title(ws, r, "運転対応の実績（HUV220低減・DEG色相改善の有無）", 4)
r = header_row(ws, r, ["操作", "期間", "HUV220低減", "DEG色相改善"])
for i, row in enumerate(C.UNTEN_JISSEKI):
    r = data_row(ws, r, list(row), [38, 16, 14, 14],
                 aligns=["left", "center", "center", "center"], alt=(i % 2 == 1))
r += 1
cell = ws.cell(r, 1, "要点：" + C.UNTEN_NOTE)
cell.font = Font(name=FONT, size=9.5, bold=True, color=INK)
cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
for c in range(1, 5):
    ws.cell(r, c).border = box
    ws.cell(r, c).fill = fillc(ALT_FILL)
ws.row_dimensions[r].height = estlines(C.UNTEN_NOTE, 90) * 15 + 8

# ============ Sheet 3: 薬剤 ============
ws = wb.create_sheet("3.薬剤")
setw(ws, [14, 98])
page(ws)
r = 1
r = title(ws, r, "NaOH（2/26〜3/30で連続投入、その後リン酸へ移行）", 2)
for k in ["投入箇所", "投入量", "効果", "機構", "留意"]:
    r = data_row(ws, r, [k, C.NAOH[k]], [12, 92], alt=False, strong_first=True)
r += 1
r = title(ws, r, "リン酸Na（Na₃PO₄・4/20〜継続中）", 2)
for k, lab in [("投入箇所", "投入箇所"), ("投入量", "投入量"), ("効果", "効果"),
               ("機構DEG", "機構（DEG側）"), ("機構MEG", "機構（MEG側）"), ("留意", "留意")]:
    r = data_row(ws, r, [lab, C.PHOS[k]], [12, 92], alt=False, strong_first=True)
r += 1
r = title(ws, r, "リン酸Na テスト履歴", 2)
setw_done = True
# テスト履歴は4列なので列幅を一時拡張
ws.column_dimensions["A"].width = 6
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 34
ws.column_dimensions["D"].width = 56
r = header_row(ws, r, ["回", "時期", "条件", "結果"])
for i, row in enumerate(C.PHOS_TEST):
    r = data_row(ws, r, list(row), [5, 14, 30, 50],
                 aligns=["center", "left", "left", "left"], alt=(i % 2 == 1))

# ============ Sheet 4: 対応方案 ============
ws = wb.create_sheet("4.対応方案")
setw(ws, [10, 30, 14, 32, 36])
page(ws)
r = 1
r = title(ws, r, "対応方案（方案1〜7）", 5)
r = header_row(ws, r, ["方案", "内容", "戦略", "ねらい", "現状"])
for i, row in enumerate(C.HOUAN):
    r = data_row(ws, r, list(row), [8, 26, 12, 28, 30],
                 aligns=["center", "left", "center", "left", "left"], alt=(i % 2 == 1), strong_first=True)
r += 1
cell = ws.cell(r, 1, "方針：" + C.HOUAN_HOSHIN)
cell.font = Font(name=FONT, size=9.5, bold=True, color=INK)
cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
for c in range(1, 6):
    ws.cell(r, c).border = box
    ws.cell(r, c).fill = fillc(ALT_FILL)
ws.row_dimensions[r].height = estlines(C.HOUAN_HOSHIN, 110) * 15 + 8

# ============ Sheet 5: 会議方針・指標 ============
ws = wb.create_sheet("5.会議方針と指標")
setw(ws, [22, 90])
page(ws)
r = 1
r = title(ws, r, "スタートアップ運転条件案（6/19 DEG対策会議）", 2)
r = header_row(ws, r, ["項目", "方針"])
for i, (a, b) in enumerate(C.SU_JOKEN):
    r = data_row(ws, r, [a, b], [20, 84], alt=(i % 2 == 1), strong_first=True)
cell = ws.cell(r, 1, C.SU_NOTE)
cell.font = Font(name=FONT, size=9, color=SUB)
cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
ws.cell(r, 1).border = box
ws.cell(r, 2).border = box
ws.row_dimensions[r].height = estlines(C.SU_NOTE, 110) * 14 + 6
r += 2
r = title(ws, r, "早期色相判断・指標", 2)
r = header_row(ws, r, ["指標", "内容"])
for i, (a, b) in enumerate(C.SHIHYO):
    r = data_row(ws, r, [a, b], [20, 84], alt=(i % 2 == 1), strong_first=True)
r += 1
r = title(ws, r, "異物・タンク・ライン処理", 2)
r = header_row(ws, r, ["対象", "内容"])
for i, (a, b) in enumerate(C.IBUTSU):
    r = data_row(ws, r, [a, b], [20, 84], alt=(i % 2 == 1), strong_first=True)
r += 1
r = title(ws, r, "顧客（レゾナック）向け説明の骨子", 2)
for i, line in enumerate(C.KOKYAKU):
    r = data_row(ws, r, ["", "・" + line], [20, 84], alt=(i % 2 == 1))
cell = ws.cell(r, 2, C.KOKYAKU_NOTE)
cell.font = Font(name=FONT, size=9, color=SUB)
cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
ws.cell(r, 1).border = box
ws.cell(r, 2).border = box
ws.row_dimensions[r].height = estlines(C.KOKYAKU_NOTE, 110) * 14 + 6

# ============ Sheet 6: 着色の筋 ============
ws = wb.create_sheet("6.着色の筋")
setw(ws, [22, 44, 48])
page(ws)
r = 1
r = title(ws, r, "着色の見立て（工程に沿った筋）", 3)
r = header_row(ws, r, ["工程", "起きていること", "根拠"])
for i, (a, b, c2) in enumerate(C.SCENARIO):
    r = data_row(ws, r, [a, b, c2], [20, 26, 28], alt=(i % 2 == 1), strong_first=True)
r += 1
cell = ws.cell(r, 1, "一本化した筋")
cell.font = Font(name=FONT, size=10, bold=True, color=INK)
cell.fill = fillc(HEAD_FILL)
cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
cell.border = box
cell2 = ws.cell(r, 2, C.SCENARIO_SUJI)
cell2.font = Font(name=FONT, size=9.5, color=INK)
cell2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
ws.cell(r, 2).border = box
ws.cell(r, 3).border = box
ws.row_dimensions[r].height = estlines(C.SCENARIO_SUJI, 90) * 15 + 8

# ============ Sheet 7: 確認事実・要検証・次アクション ============
ws = wb.create_sheet("7.事実と次アクション")
setw(ws, [8, 104])
page(ws)
r = 1
r = title(ws, r, "確認された事実", 2)
for i, (k, v) in enumerate(C.FACTS):
    r = data_row(ws, r, [k, v], [6, 100], alt=(i % 2 == 1), strong_first=True)
r += 1
r = title(ws, r, "重要な更新（6/17〜6/20）", 2)
for i, (k, v) in enumerate(C.UPDATES):
    r = data_row(ws, r, [k, v], [6, 100], alt=(i % 2 == 1), strong_first=True)
r += 1
r = title(ws, r, "次アクション", 2)
for i, v in enumerate(C.NEXT, 1):
    r = data_row(ws, r, [str(i), v], [6, 100], alt=(i % 2 == 0))
r += 1
r = title(ws, r, "要検証・齟齬", 2)
for i, v in enumerate(C.OPEN, 1):
    r = data_row(ws, r, [str(i), v], [6, 100], alt=(i % 2 == 0))

# ---- メタデータ（作成者：平尾名義／AI痕跡なし）----
wb.properties.creator = C.AUTHOR
wb.properties.lastModifiedBy = C.AUTHOR
wb.properties.title = "EG品質の現状とDEG色相悪化への対応"
wb.properties.description = ""
wb.properties.keywords = ""

path = os.path.join(OUT, "EG品質_DEG色相_事実整理.xlsx")
wb.save(path)
print("saved:", path)
