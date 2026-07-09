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

ISSUE = ("イシュー（草案）：色相悪化は全タンクで起きているが、T-604がNaOH停止後の液・リン酸添加品を受けても"
         "APHA5未満から30まで上がり続けた（T-615は低下傾向）のは、液の質だけでは説明できず、"
         "むしろT-604固有の条件（TEGからの品名変更で4月に開放復帰した内面状態・コーティング・"
         "3基中最小の500tゆえの壁面/底板影響）が悪化を加速しているのでは。")

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
    ("A", "T-604固有の条件（内面・コーティング・幾何）が、全タンク共通の悪化を加速しているのでは", "", "", 0),
    ("A-1", "コーティング仕様の食い違い（本日会議発言＝元々FRP／4/1調査資料＝エポキシ）をまず決着させるべきはず", "", "", 1),
    ("A-1-①", "図面・塗装仕様書・補修履歴で仕様と補修歴を確定できるはず",
     "操油課・機械Grに図面・塗装記録・膜厚データの確認を依頼（過去の補修履歴含む）",
     "FRPなら他タンクとの仕様差が主対象に。エポキシなら材質差仮説は落ちて内面状態（劣化・剥離）の議論に移る", 2),
    ("A-1-②", "FRPだった場合、DEG・有機酸への耐性がエポキシより弱く、劣化・溶出・剥離が「加速」の説明になるのでは",
     "剥離片・付着物の分析（樹脂成分の同定）。膜厚・剥離範囲の記録",
     "再コーティング仕様（エポキシ化・全面補修）を復旧計画に織り込む", 2),
    ("A-2", "TEGからの品名変更（4月開放復帰）の影響＝TEG時代の残渣・付着や、開放時の清掃・点検の程度が効いたのでは", "", "", 1),
    ("A-2-①", "4月開放時の点検・清掃の記録と今回所見を照合すれば、当時からあったものか今回できたものか分かるはず",
     "4月開放時の点検記録との照合（機械Gr宿題・石塚さん論点＝前回復旧判断との整合）",
     "点検・復旧判断の基準を明文化し、今回の復旧判断に織り込む", 2),
    ("A-2-②", "4月開放時に四隅から採取した残液と今回の残液・付着物に、タンク由来成分（Fe・コーティング成分・TEG由来）が出るはず",
     "四隅残液サンプルの分析（野田さん宿題・金枝さん受け渡し調整）。T-555（茶褐色汚れ）・T-615（鉄サビ・砂＋鉄10〜12%）の分析結果と比較",
     "タンク由来成分が出れば清掃・内面処理を色相対策に昇格。出なければ液側（B系）が強まる", 2),
    ("A-3", "3基中最小（T-604 500t・T-615 950kL・T-555 2,000kL）で壁面/液比・底板影響（デッド在庫約2t・底板とドレンノズル間25mm）が最も濃く出るのでは",
     "APHA上昇速度を在庫量・接液面積で規格化してタンク間比較",
     "規格化で差が消えるなら「タンク固有」でなく共通機構＋幾何の差。消えなければA-1/A-2の固有要因を追う", 1),
    ("A-4", "底板の状態不良・「他タンクより加速」（石塚さん所感）はA-1〜A-3の帰結として説明できるはず",
     "T-604開放検査記録（写真・堆積物性状・コーティング状態・肉厚）の共有依頼",
     "劣化速度が速ければ補修範囲・再発防止（内面仕様・受入液の管理）を復旧計画に反映", 1),
    ("B", "（対抗仮説）液側で説明できる余地＝Na停止後でも残留Na・高A-ALD分が入っていたのでは", "", "", 0),
    ("B-1", "塔内Na濃度は「低下傾向」であってゼロではなく、4/10〜4/20受入分はリン酸開始前で A-ALDも高かったはず",
     "4月受入分の品質実績（Na・pH・A-ALD）確認。DEG色相UVデータ整形版の並記",
     "液側で説明が付くなら、S/U後の受入管理（前駆体を含む液の仕分け）で対応", 1),
    ("B-2", "全タンク悪化（5/11整理「タンク固有の可能性低い」）とT-604の速度差は、共通機構＋タンク条件の重ね合わせで説明できるはず",
     "タンク別のAPHA・UV450上昇速度の横並び比較（リン酸添加品受入の前後で区切る）",
     "共通機構側の対策（前駆体を入れない）とタンク側の対策（清掃・内面）の効き分けを整理", 1),
    ("C", "復旧判断（27日頃見込み）までに、色相と切り離して決めるべき設備側の判断があるはず", "", "", 0),
    ("C-1", "前回（4月）復旧判断との整合＝「前回ちゃんと点検できていないで復旧した」と言われない整理が要るはず",
     "4月開放時の記録確認（機械Gr宿題）",
     "点検・復旧の判断基準を記録に残す", 1),
    ("C-2", "復帰後のT-604の使い方（端切り・共洗いの受け皿とするか＝6/22相談事項）は、原因の向きが決まってから決めるべきはず",
     "A系・B系の見極め結果を待って判断材料を揃える",
     "受け皿とする場合は悪化品の隔離運用、しない場合はS/U共洗い計画の見直し", 1),
]

# 背骨（事実・見極め・対応）
BACKBONE = [
    ("事実", "T-604はTEGタンクとして運用後（2025年12月〜2026年4月中旬）、品名変更（TEG→DEG）のため3〜4月に開放し、4/10からDEG受入を開始（色相5未満）。"),
    ("事実", "受入品はNaOH停止（3/30）後の生産DEGが主で、4/20からはリン酸添加品。それでもAPHAは4/20の5未満から5/6に15（規格超過）、5/13に20、5/18に30まで悪化（実測）。"),
    ("事実", "色相悪化は全タンクで発生（5/11調査の整理＝「タンク固有の不具合の可能性は低い」）。ストリーム品は常に規格内で、タンク（N2封入）保管中にのみ上昇。T-615はリン酸添加後に低下傾向、T-604は上昇継続＝速度差が特徴。"),
    ("事実", "7/9開放（7/6〜清掃着手）でT-555・T-615と比較して状態が良くない・底板の状態が良くない・タッチアップ箇所が想定より多そう（金枝・木村）。所見詳細の記録は未入手。"),
    ("事実（食い違い）", "コーティング仕様＝本日会議発言では「T-604は元々FRP・タッチアップはエポキシ・T-555/T-615はエポキシ仕様」。一方4/1調査資料では「T-604内壁＝エポキシ樹脂コーティング（耐薬品性、実績あり）」。要決着。"),
    ("事実", "3基構成＝T-555 2,000kL・T-615 950kL・T-604 500t（最小）。T-604のデッド在庫約2t（底板とドレンノズル間25mm）。N2シール・シールポット・サンプリング装置は5/11時点で異常なし。"),
    ("既存整理", "着色物質はポリエナール（GPC実測）。鉄錆単独では着色しない（保管試験）。前駆体はタンク長滞留で鎖伸長し、使い切られるとAPHA30前後で安定（T-555残液実測）。"),
    ("見極め", "①コーティング仕様の決着（図面・塗装記録・補修履歴）＝A-1。"),
    ("見極め", "②4月開放時の点検・清掃記録と今回所見の照合（前回復旧判断の整合）＝A-2-①。"),
    ("見極め", "③残液・付着物分析（四隅サンプル含む）でタンク由来成分（Fe・コーティング成分・TEG由来）が出るか＝A-2-②。"),
    ("見極め", "④上昇速度を在庫量・接液面積で規格化した横並び比較＝A-3・B-2。"),
    ("対応", "タンク側が当たりなら清掃・再コーティング仕様と復帰後の使い方（共洗い受け皿の是非）を決める。液側の余地（残留Na・4/10〜4/20の高A-ALD分）はB系で並行確認。どちらに転んでも次の打ち手が決まる。"),
]

PENDING = [
    "本日（7/9）会議議事録＝開放所見の詳細（底板・側板・コーティングの状態、汚れ・錆の性状・範囲）",
    "T-604開放検査記録（写真・堆積物性状・コーティング状態・肉厚測定）",
    "コーティング仕様の確定資料（図面・塗装仕様書・補修履歴・膜厚）",
    "4月開放時の点検記録（機械Gr宿題）",
    "四隅残液サンプルの分析結果（野田さん宿題）",
    "タンク別APHA・UV450上昇速度の横並びデータ（在庫量・接液面積での規格化用）",
]

AGREE = [
    "イシューを「全タンク悪化の中でのT-604の速度差＝タンク固有条件の重ね合わせ」と設定し直した向きでよいか",
    "コーティング仕様の食い違い（会議発言FRP／4/1資料エポキシ）の決着を誰に確認するか（操油課図面か機械Grか）",
    "「検証方法」列を何に向けた検証として書くか（原因究明用か、会議説明用か）",
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
