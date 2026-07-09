# -*- coding: utf-8 -*-
"""T-604 DEG色相悪化：課題ツリーたたき台（課題設定と分離.sop準拠）Excel生成。

シート1＝課題ツリー（列は「課題（仮説）／検証方法／対応（検証が当たれば）」。
検証・対応は手を動かせる葉にだけ書き、上位ノードは空白）。
シート2＝背骨（事実・見極め・対応）と情報待ち。
配色はモノクロ（白地・黒文字・細罫線・見出しは薄グレー）＝DEG着色_フロー整理.xlsxと同体裁。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs")
os.makedirs(OUT, exist_ok=True)

DATE = "2026-07-09"
TITLE = "T-604 DEG色相悪化　課題ツリー（たたき台・合意待ち草案）"

ISSUE = ("イシュー（草案）：T-604だけAPHAが上がり続けるのは、タンク固有の異常（内面の錆・付着が着色を促進）"
         "ではなく、むしろ貯めた液の履歴（NaOH停止後に初めて貯め始めた初期液が前駆体を多く含み、"
         "使い切られるまで鎖伸長が続いている）のせいなのでは。")

HEADER_FILL = "E6E6E6"
LABEL_FILL = "F2F2F2"
TOP_FILL = "EFEFEF"
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
    for para in text.split("\n"):
        n += max(1, -(-len(para) // cpl))
    return n


# (番号, 課題（仮説）, 検証方法, 対応（検証が当たれば）, 階層0/1/2)
TREE = [
    ("A", "T-604だけAPHAが上がり続けるのは、貯めた液の履歴のせいなのでは（タンク固有の異常ではなく）", "", "", 0),
    ("A-1", "初期在庫（NaOH停止後からリン酸開始前）は前駆体を最も多く含む液だったのでは", "", "", 1),
    ("A-1-①", "貯め始めの時期は悪化品期（3月末から4/20頃）に重なるはず",
     "受払履歴で貯め始め日・受入元を確認",
     "前駆体を含む液が残る間は仕分け・ブレンド前提の出荷運用", 2),
    ("A-1-②", "T-604残液はT-615残液より前駆体が多いはず",
     "残液のUVチャート・GPCを両タンクで比較",
     "残液の処理（抜き出し・希釈）を検討", 2),
    ("A-2", "リン酸添加品を後から入れてもAPHAが上がり続けたのは、タンク内で鎖伸長が続いているためなのでは", "", "", 1),
    ("A-2-①", "上昇は前駆体が使い切られるまで続き、その後APHA30前後で頭打ちになるはず",
     "T-604のAPHA・UV450推移が飽和型か（T-555の5/18以降安定と同型か）を確認",
     "頭打ちが確認できれば、在庫の作り直しでなく時間経過とブレンドで収束を待つ選択肢が立つ", 2),
    ("A-2-②", "リン酸は上流の前駆体生成に効き、タンク内の反応には効かないのでは",
     "T-615（低下）との上昇速度比較で切り分け",
     "リン酸投入判断を「タンクを治す」でなく「前駆体を入れない」目的に整理し直す", 2),
    ("A-3", "（反証側）同じ液を入れてもT-604だけ上がるなら、タンク固有差なのでは",
     "同等液の入槽先別UV450比較。ボンベサンプリング（プラント由来かタンク由来かの切り分け）",
     "タンク固有差ならB系（設備側）を色相対策に昇格", 1),
    ("B", "開放で内面の状態が良くないのは、着色の原因ではなくむしろ結果（初期液の有機酸による腐食）なのでは", "", "", 0),
    ("B-1", "NaOH停止で有機酸（ギ酸）の中和がなくなり、初期液が内面を腐食させたのでは",
     "残液の有機酸・pH実測。腐食部位の分布（液相部か気相部か）の確認",
     "腐食は液側の運用改善（前駆体・有機酸を入れない）で再発防止", 1),
    ("B-2", "（反証側）付着物が450nm成長の触媒・供給源として働いているのでは",
     "付着物の元素分析（Fe・P・Na）とDEG共存保管試験で450nm再現の有無（T-555異物「ほぼ鉄」と比較）",
     "再現すればタンク清掃・内面処理を色相対策に昇格", 1),
    ("B-3", "腐食の程度が使用継続に耐えるかは、色相と切り離した設備健全性の判断のはず",
     "肉厚測定・開放検査結果の整理",
     "補修・使用可否の判断（色相の結論を待たずに進める）", 1),
]

# 背骨（事実・見極め・対応）
BACKBONE = [
    ("事実", "T-604はNaOH停止後に初めて貯め始めたタンク（それまで在庫なし）。"),
    ("事実", "リン酸添加品を貯めてもT-604はAPHA上昇が継続。T-615はリン酸添加後に低下傾向＝この実測差が最大の特徴。"),
    ("事実", "T-604開放で内面の状態が良くない（所見の詳細は情報収集中）。T-555（6/16初回開放）は比較的綺麗・異物はほぼ鉄。T-615（6/17）もT-555と同程度。"),
    ("事実", "既存整理＝前駆体はタンク長滞留で鎖伸長し、使い切られるとAPHA30前後で安定（T-555残液実測）。鉄錆単独では着色しない（保管試験）。主要因はA-ALD濃度上昇とNa存在の組み合わせ（7/1ほぼ合意）。"),
    ("見極め", "①T-604とT-615の残液で前駆体（UVチャート・GPC）に差があるか（A-1-②）。"),
    ("見極め", "②付着物が450nm成長の触媒になるか＝錆・付着が原因か結果かの向き決め（B-2）。"),
    ("見極め", "③T-604のAPHA上昇が飽和型か上がり続けているか（A-2-①）。"),
    ("対応", "液の履歴が当たりなら運用側（前駆体を含む液の受入管理・仕分け・ブレンド・残液処理）。反証が当たりなら設備側（清掃・補修・使用可否）。どちらに転んでも次の打ち手が決まる。"),
]

PENDING = [
    "T-604開放所見の詳細（部位・色・付着量・腐食の程度・写真）",
    "タンク受払履歴（貯め始め日・受入元・悪化品期との重なり・ブレンド比率）",
    "残液・付着物の分析結果（UVチャート・GPC・pH・Na・Fe・P・有機酸）",
    "T-604とT-615のAPHA・UV450推移の並記データ（リン酸添加品受入の前後）",
]

AGREE = [
    "イシュー1文の向き（「タンク固有ではなく液の履歴」という対置）でよいか",
    "「検証方法」列を何に向けた検証として書くか（原因究明用か、会議説明用か）",
    "A-2-①「頭打ちになるはず」は足元のT-604推移次第で書き方を変える（上がり続けているなら弱い主張）",
]


wb = Workbook()
ws = wb.active
ws.title = "課題ツリー"
ws.sheet_view.showGridLines = False

COLS = [("番号", 9), ("課題（仮説）", 52), ("検証方法", 36), ("対応（検証が当たれば）", 36)]
for i, (_, w) in enumerate(COLS):
    ws.column_dimensions[get_column_letter(1 + i)].width = w
last_letter = get_column_letter(len(COLS))
CPLS = [int(w / 2.05) for _, w in COLS]

r = 1
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, TITLE)
c.font = Font(name=FONT, size=15, bold=True, color=INK)
c.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[r].height = 26
r += 1
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, ISSUE)
c.font = Font(name=FONT, size=10, bold=True, color=INK)
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws.row_dimensions[r].height = est_lines(ISSUE, sum(CPLS)) * 15 + 12
r += 1
ws.merge_cells(f"A{r}:{last_letter}{r}")
c = ws.cell(r, 1, f"（{DATE} 時点・たたき台）")
c.font = Font(name=FONT, size=9, color=SUB_INK)
c.alignment = Alignment(horizontal="right", vertical="center")
for cidx in range(1, len(COLS) + 1):
    ws.cell(r, cidx).border = Border(bottom=med)
ws.row_dimensions[r].height = 14
r += 1

HEADER_ROW = r
for i, (name, _) in enumerate(COLS):
    cell = ws.cell(r, 1 + i, name)
    cell.font = Font(name=FONT, size=10, bold=True, color=INK)
    cell.fill = fill(HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=med)
ws.row_dimensions[r].height = 22
r += 1

for num, hypo, verify, action, level in TREE:
    indent = level
    row_fill = TOP_FILL if level == 0 else WHITE
    bold = level == 0
    vals = [num, hypo, verify, action]
    maxlines = 1
    for i, txt in enumerate(vals):
        cell = ws.cell(r, 1 + i, txt)
        cell.font = Font(name=FONT, size=9, bold=bold if i <= 1 else False,
                         color=INK if i <= 1 else SUB_INK)
        cell.fill = fill(LABEL_FILL if i == 0 else row_fill)
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True,
                                   indent=(indent + 1 if i == 1 else 1))
        cell.border = border
        maxlines = max(maxlines, est_lines(txt, max(CPLS[i] - (indent if i == 1 else 0), 8)))
    ws.row_dimensions[r].height = maxlines * 12.5 + 7
    r += 1

ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.4
ws.freeze_panes = "A" + str(HEADER_ROW + 1)


ws2 = wb.create_sheet("背骨と情報待ち")
ws2.sheet_view.showGridLines = False
ws2.column_dimensions["A"].width = 12
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


def s2_kv(r, k, v):
    kc = ws2.cell(r, 1, k)
    kc.font = Font(name=FONT, size=9.5, bold=True, color=INK)
    kc.fill = fill(LABEL_FILL)
    kc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    kc.border = border
    vc = ws2.cell(r, 2, v)
    vc.font = Font(name=FONT, size=9.5, color=INK)
    vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    vc.border = border
    ws2.row_dimensions[r].height = max(26, est_lines(v, 52) * 15 + 7)
    return r + 1


r = 1
r = s2_band(r, "背骨（事実・見極め・対応）")
for k, v in BACKBONE:
    r = s2_kv(r, k, v)
r += 1
r = s2_band(r, "情報待ち（入り次第ツリーに反映）")
for i, v in enumerate(PENDING, 1):
    r = s2_kv(r, str(i), v)
r += 1
r = s2_band(r, "合意待ちの点（この草案で確認すること）")
for i, v in enumerate(AGREE, 1):
    r = s2_kv(r, str(i), v)

ws2.page_setup.orientation = "landscape"
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 0
ws2.sheet_properties.pageSetUpPr.fitToPage = True

wb.properties.creator = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.lastModifiedBy = "hirao kazuaki/0465811/平尾　一陽"
wb.properties.title = TITLE
wb.properties.description = ""
wb.properties.keywords = ""

path = os.path.join(OUT, "T-604_DEG色相_課題分離.xlsx")
wb.save(path)
print("saved:", path)
