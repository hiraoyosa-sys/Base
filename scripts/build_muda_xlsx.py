# -*- coding: utf-8 -*-
"""業務効率化／無駄の見える化 Excel生成。

シート構成：
  1) 表紙・中心命題・定義・スコープ
  2) 再現性の構成4要件
  3) 見える化マップ（状態主語＋優先度）
  4) 診断チェックリスト
  5) 出口戦略・人への対応・ロジック・最初の一手・発信スコープ
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）で content.py 系の資料と統一。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import content_muda as C

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)

HEADER_FILL = "E6E6E6"
LABEL_FILL = "F2F2F2"
SUB_FILL = "F7F7F7"
WHITE = "FFFFFF"
INK = "000000"
SUB_INK = "404040"
BORDER_CLR = "808080"

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


# ---------- 共通ヘルパ ----------
def title_block(ws, title, sub, last_letter, last_col):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(f"A1:{last_letter}1")
    c = ws.cell(1, 1, title)
    c.font = Font(name=FONT, size=15, bold=True, color=INK)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells(f"A2:{last_letter}2")
    c = ws.cell(2, 1, sub)
    c.font = Font(name=FONT, size=9.5, color=SUB_INK)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for cidx in range(1, last_col + 1):
        ws.cell(2, cidx).border = Border(bottom=med)
    ws.row_dimensions[2].height = 30
    c = ws.cell(3, 1, f"（{C.DATE} 時点・木村さん 6/17発信への対応）")
    c.font = Font(name=FONT, size=9, color=SUB_INK)
    return 4


def band(ws, r, text, span_letter):
    ws.merge_cells(f"A{r}:{span_letter}{r}")
    c = ws.cell(r, 1, text)
    c.font = Font(name=FONT, size=11.5, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, ord(span_letter) - 64 + 1):
        ws.cell(r, col).border = border
    ws.row_dimensions[r].height = 22
    return r + 1


def para(ws, r, text, span_letter, cpl=110, size=10, bold=False, ink=INK):
    ws.merge_cells(f"A{r}:{span_letter}{r}")
    c = ws.cell(r, 1, text)
    c.font = Font(name=FONT, size=size, bold=bold, color=ink)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    for col in range(1, ord(span_letter) - 64 + 1):
        ws.cell(r, col).border = border
    ws.row_dimensions[r].height = est_lines(text, cpl) * 15 + 8
    return r + 1


# ========== シート1：表紙・中心命題・定義 ==========
ws = wb.active
ws.title = "命題と定義"
ws.column_dimensions["A"].width = 20
for col in "BCDE":
    ws.column_dimensions[col].width = 30
r = title_block(ws, C.TITLE, C.SUBTITLE, "E", 5)
r += 1
r = band(ws, r, "中心命題 ── 無駄とは“労力を二度払う状態”", "E")
r = para(ws, r, C.THESIS, "E", cpl=108)
r += 1
r = band(ws, r, "この資料のスコープ（何を扱い、何を扱わないか）", "E")
r = para(ws, r, C.SCOPE, "E", cpl=108, ink=SUB_INK)
r += 1
r = band(ws, r, "定義パネル", "E")
# 定義テーブル：A=用語, B:E=定義
for term, d in C.DEFS:
    tc = ws.cell(r, 1, term)
    tc.font = Font(name=FONT, size=10, bold=True, color=INK)
    tc.fill = fill(LABEL_FILL)
    tc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    tc.border = border
    ws.merge_cells(f"B{r}:E{r}")
    dc = ws.cell(r, 2, d)
    dc.font = Font(name=FONT, size=9.5, color=INK)
    dc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for col in range(1, 6):
        ws.cell(r, col).border = border
    ws.row_dimensions[r].height = est_lines(d, 88) * 15 + 8
    r += 1

# ========== シート2：再現性の構成4要件 ==========
ws2 = wb.create_sheet("再現性4要件")
ws2.sheet_view.showGridLines = False
cols2 = [("A", 14), ("B", 22), ("C", 34), ("D", 40), ("E", 40)]
for col, w in cols2:
    ws2.column_dimensions[col].width = w
r = title_block(ws2, "再現性の構成4要件", "いずれの欠落も“二度払い”を生む。分類ではなく、再現性を成り立たせる要件。", "E", 5)
heads = ["要件", "一言でいうと", "欠落すると（どの二度払い）", "自分の業務での現れ（構造の層）", "現場・朝会での現れ"]
for i, h in enumerate(heads):
    c = ws2.cell(r, 1 + i, h)
    c.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=med)
ws2.row_dimensions[r].height = 26
r += 1
cpls = [10, 16, 26, 30, 30]
for req, one, lack, mine, field in C.REQS:
    vals = [req, one, lack, mine, field]
    maxlines = 1
    for i, v in enumerate(vals):
        c = ws2.cell(r, 1 + i, v)
        bold = (i == 0)
        c.font = Font(name=FONT, size=9, bold=bold, color=INK if i <= 2 else SUB_INK)
        c.fill = fill(LABEL_FILL if i == 0 else WHITE)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        c.border = border
        maxlines = max(maxlines, est_lines(v, cpls[i]))
    ws2.row_dimensions[r].height = maxlines * 13 + 7
    r += 1

# ========== シート3：見える化マップ ==========
ws3 = wb.create_sheet("見える化マップ")
ws3.sheet_view.showGridLines = False
w3 = [("A", 46), ("B", 10), ("C", 8), ("D", 8), ("E", 8), ("F", 44)]
for col, w in w3:
    ws3.column_dimensions[col].width = w
r = title_block(ws3, "無駄の見える化マップ", "主語は“業務”でなく“状態”。優先度＝規模×容易性で着手対象を絞る。", "F", 6)
for i, h in enumerate(C.MAP_COLS):
    c = ws3.cell(r, 1 + i, h)
    c.font = Font(name=FONT, size=9, bold=True, color=INK)
    c.fill = fill(HEADER_FILL)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(left=thin, right=thin, top=thin, bottom=med)
ws3.row_dimensions[r].height = 30
r += 1
cpls3 = [42, 8, 6, 6, 6, 40]
for row in C.MAP_ROWS:
    maxlines = 1
    for i, v in enumerate(row):
        c = ws3.cell(r, 1 + i, v)
        c.font = Font(name=FONT, size=9, color=INK if i in (0, 5) else SUB_INK)
        align_h = "center" if i in (1, 2, 3, 4) else "left"
        c.alignment = Alignment(horizontal=align_h, vertical="top", wrap_text=True, indent=1 if align_h == "left" else 0)
        c.fill = fill(SUB_FILL if str(row[1]) == "—" else WHITE)
        c.border = border
        maxlines = max(maxlines, est_lines(v, cpls3[i]))
    ws3.row_dimensions[r].height = maxlines * 13 + 8
    r += 1
r += 1
r = para(ws3, r, "最初の1アクション：" + C.FIRST_ACTION, "F", cpl=128, bold=True)

# ========== シート4：診断チェックリスト ==========
ws4 = wb.create_sheet("診断チェックリスト")
ws4.sheet_view.showGridLines = False
w4 = [("A", 10), ("B", 60), ("C", 50)]
for col, w in w4:
    ws4.column_dimensions[col].width = w
r = title_block(ws4, "診断チェックリスト", "「無駄か？」でなく「再現性があるか？」を問う。Noの理由＝改善の本丸。", "C", 3)
for gname, items in C.CHECK_GROUPS:
    r = band(ws4, r, gname, "C")
    for name, q, no in items:
        nc = ws4.cell(r, 1, name)
        nc.font = Font(name=FONT, size=9.5, bold=True, color=INK)
        nc.fill = fill(LABEL_FILL)
        nc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        nc.border = border
        qc = ws4.cell(r, 2, q)
        qc.font = Font(name=FONT, size=9.5, color=INK)
        qc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        qc.border = border
        oc = ws4.cell(r, 3, "Noなら→ " + no)
        oc.font = Font(name=FONT, size=9, color=SUB_INK)
        oc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        oc.border = border
        ws4.row_dimensions[r].height = max(est_lines(q, 56), est_lines(no, 46)) * 14 + 8
        r += 1

# ========== シート5：出口戦略・人への対応・ロジック・発信 ==========
ws5 = wb.create_sheet("出口戦略と人への対応")
ws5.sheet_view.showGridLines = False
ws5.column_dimensions["A"].width = 24
for col in "BCD":
    ws5.column_dimensions[col].width = 30
r = title_block(ws5, "出口戦略・人への対応・上位目的", "あるべき状態から逆算（バックキャスト）。打ち手＝器は最後。", "D", 4)
r = band(ws5, r, "出口戦略（バックキャスティング）", "D")
r = para(ws5, r, C.EXIT_GOAL, "D", cpl=92, bold=True)
for step, desc in C.EXIT_STEPS:
    sc = ws5.cell(r, 1, step)
    sc.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    sc.fill = fill(LABEL_FILL)
    sc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    sc.border = border
    ws5.merge_cells(f"B{r}:D{r}")
    dc = ws5.cell(r, 2, desc)
    dc.font = Font(name=FONT, size=9.5, color=INK)
    dc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for col in range(1, 5):
        ws5.cell(r, col).border = border
    ws5.row_dimensions[r].height = est_lines(desc, 82) * 15 + 8
    r += 1
r = para(ws5, r, C.EXIT_TOOL, "D", cpl=92, ink=SUB_INK)
r += 1
r = band(ws5, r, "人への対応（情報分断・属人化の関係者マネジメント）", "D")
r = para(ws5, r, C.PEOPLE_PRINCIPLE, "D", cpl=92, ink=SUB_INK)
for who, how in C.PEOPLE_ROWS:
    wc = ws5.cell(r, 1, who)
    wc.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    wc.fill = fill(LABEL_FILL)
    wc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    wc.border = border
    ws5.merge_cells(f"B{r}:D{r}")
    hc = ws5.cell(r, 2, how)
    hc.font = Font(name=FONT, size=9.5, color=INK)
    hc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for col in range(1, 5):
        ws5.cell(r, col).border = border
    ws5.row_dimensions[r].height = est_lines(how, 82) * 15 + 8
    r += 1
r += 1
r = band(ws5, r, "上位目的への接続（残業削減は結果指標）", "D")
r = para(ws5, r, "　→　".join(C.LOGIC), "D", cpl=92)
r += 1
r = band(ws5, r, "今日の発信スコープ案", "D")
for name, desc in C.RELEASE:
    nc = ws5.cell(r, 1, name)
    nc.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    nc.fill = fill(LABEL_FILL)
    nc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    nc.border = border
    ws5.merge_cells(f"B{r}:D{r}")
    dc = ws5.cell(r, 2, desc)
    dc.font = Font(name=FONT, size=9.5, color=INK)
    dc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for col in range(1, 5):
        ws5.cell(r, col).border = border
    ws5.row_dimensions[r].height = est_lines(desc, 82) * 15 + 8
    r += 1

# ---- 全シート共通：印刷設定 ----
for sh in wb.worksheets:
    sh.page_setup.orientation = "landscape"
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.page_margins.left = sh.page_margins.right = 0.3
    sh.page_margins.top = sh.page_margins.bottom = 0.4

# ---- メタデータ（作成者）----
wb.properties.creator = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.lastModifiedBy = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.title = "業務効率化／無駄の見える化"

path = os.path.join(OUT, "無駄の見える化_再現性で見る.xlsx")
wb.save(path)
print("saved:", path)
