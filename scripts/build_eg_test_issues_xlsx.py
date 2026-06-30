# -*- coding: utf-8 -*-
"""EG品質確認テスト(EG-1/EG-2) 課題整理 Excel生成（本筋に絞った見やすい版）。

出力 : outputs/EG品質確認テスト_課題整理_20260630.xlsx
方針 : 本筋に厳選し各セルを短文化。先頭に「テスト概要」を置き全体像を1画面で把握。
シート: ①テスト概要 ②課題一覧 ③RD事前提出依頼 ④スケジュール
体裁 : モノクロ・マス結合なし・装飾/造語/矢印記号なし・出典/作成者表記なし。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "outputs")
os.makedirs(OUT, exist_ok=True)

FONT = "游ゴシック"
HEADER_FILL = "D9D9D9"
GROUP_FILL = "F2F2F2"
INK = "000000"
BORDER_CLR = "808080"
thin = Side(style="thin", color=BORDER_CLR)
border = Border(left=thin, right=thin, top=thin, bottom=thin)


def fill(c):
    return PatternFill("solid", fgColor=c)


# ====== テスト概要（本筋を1画面で）======
OVERVIEW = [
    ("目的", "M50の品質影響を実機で確認する。プラントを止めて悪化を避けるのでなく、品質対策を効かせた状態で規格内に収まるかを確かめる。"),
    ("テスト構成", "2段階。EG-1=低生産量(EO稼働50%・MEG約7.3t/h近傍)。EG-2=低水比(原料水を下げEOを増やし稼働70%近くへ、反応器入口ALDを将来想定水準へ)。"),
    ("時期", "7月末〜8月頭の2週間弱。EG-1=7/28〜8/2、EG-2=8/3〜8/5（前提により変動）。"),
    ("織込むカウンター", "リン酸(リン酸Na)添加、MEG塔頂圧ダウン(27から23kPaA、塔底167℃以下)、温度低下。テスト開始時から併用する。"),
    ("中止判断", "製品MEG HUV220≦0.15。客先のHUV220は自社測定不可のため管理値で代用。"),
    ("最大の懸念", "反応器入口の不純物濃縮（循環水A-ALDが現状11から約33ppm＝過去実績の外側）。AALDは下げにくく、出口戦略(ガード)が前提。"),
    ("明日の品質定例で必要", "RDに事前提出を依頼する事項（③RD事前提出依頼シート）。中核はリン酸だけでDEG色相を保てる根拠データ。"),
    ("当面の関門", "事業部(アルB)との個別議論、各SA(課1次・部1次・課3次)の起票、SU前準備、テスト、M50の8/20ゲートの順で進む。"),
]

# ====== 課題一覧（本筋・約29件）======
# 分類順: 運転, 品質, 設備, 会議, 推進
MASTER = [
    # --- 運転 / プロセス ---
    dict(b="運転/プロセス", k="EO稼働50%の達成手段（強制循環の継続運用）", s="50%は強制循環が必要。反応器温度50℃維持と両立要", h="立上げ時の強制循環の継続を含め製造と運用を相談。変更管理事項", g="テスト前", t="木村/製造", rd="", y="高"),
    dict(b="運転/プロセス", k="EG-1（低生産量）の手順確定", s="EO稼働50%・MEG7.3t/h近傍・約6日", h="EOを絞り段階的にロード調整し品質を確認", g="7/28〜8/2", t="平尾/木村/製造", rd="", y="高"),
    dict(b="運転/プロセス", k="EG-2（低水比）の手順確定", s="原料水を下げEO増・稼働70%近く・約3日", h="反応器入口ALDを将来想定水準へ上げ品質を確認", g="8/3〜8/5", t="平尾/木村/製造", rd="", y="高"),
    dict(b="運転/プロセス", k="テスト時EO稼働率の確定（50%）", s="60%だと目標7.3に届かない", h="50%で確定（過去案の60%は更新前）", g="計画確定時", t="平尾/木村", rd="", y="中"),
    # --- 運転 / タンク繰り ---
    dict(b="運転/タンク繰り", k="外販EOタンク繰りの成立", s="在庫ピーク約6,293t・お盆。外販EOタンク2基使用不可", h="中間タンクのバッファ運用と引取増で成立点を確保", g="7/5前後", t="本田/遠藤/物流", rd="", y="高"),
    dict(b="運転/タンク繰り", k="DEGタンク組分けと東レ向け出荷不可期間", s="期間中の仕切り要。3期DEGタンクの戻り時期が未確認", h="出荷不可期間を具体化し関係者へ共有。戻り時期を確認", g="テスト期間中", t="木村/遠藤/物流品証", rd="", y="高"),
    dict(b="運転/タンク繰り", k="製販バランス（生産量を絞る）", s="製品タンク収容・外販やPEGは後回し", h="日別バランス表で在庫を管理しながら絞る", g="計画確定時", t="本田/木村", rd="", y="中"),
    # --- 品質 / RD検証（明日の依頼の核）---
    dict(b="品質/RD検証", k="リン酸(リン酸Na)だけでDEG色相が保てる根拠データ", s="はっきりしたデータが無い。SAで提示必須", h="あり・なし比較で根拠化し、SAで示せる形に整理", g="SU前/SA前", t="基礎研/木村", rd="○", y="高"),
    dict(b="品質/RD検証", k="M50想定のFALD・AALD推定根拠データ", s="AALDが支配的。経験外の濃度域への外挿", h="推定根拠を提示できる形に整理", g="SA前", t="木村/基礎研", rd="○", y="高"),
    dict(b="品質/RD検証", k="色相悪化メカニズム仮説の説明ストラテジー", s="ベンチ未再現で断定困難", h="事実・仮説・対応で1枚に整理し基礎研と擦り合わせ", g="SU前", t="基礎研/木村", rd="○", y="高"),
    dict(b="品質/RD検証", k="出口戦略(IER/吸着)でのアルデヒド挙動データ", s="樹脂でアルデヒドも増える裏付けあり", h="挙動データを整理しガード設計に反映", g="SA前", t="基礎研/木村", rd="○", y="中"),
    dict(b="品質/RD検証", k="方案3（FALDパージ）の有効性と織込み要否", s="期待薄との意見。AALD上限の見極めが本筋", h="有効性を整理し織込み要否を基礎研と決定", g="8/20ゲート前", t="木村/髙橋", rd="○", y="中"),
    dict(b="品質/RD検証", k="EG-2低水比での不純物の濃縮/希釈の向き", s="濃縮で品質影響を加速する可能性", h="向きをRDで確認（どちらでも低水比は切り分け前提）", g="EG-2前", t="基礎研/木村", rd="○", y="中"),
    # --- 品質 / 対策（テスト織込み）---
    dict(b="品質/対策", k="カウンター3点の織込みと手順明記", s="塔底167℃以下・G-ALD≦2,500g/h", h="リン酸添加・塔頂圧ダウン・温度低下をテスト開始時から併用", g="7/27週", t="平尾/木村/製造", rd="", y="高"),
    dict(b="品質/対策", k="中止判断基準（MEG HUV220≦0.15）の運用", s="客先HUV220は自社測定不可", h="管理値0.15で代用し合否判定に用いる", g="テスト期間中", t="平尾/木村", rd="", y="高"),
    dict(b="品質/対策", k="結果の出し方（2〜3点刻みで7.3外挿）", s="将来条件は実機で完全には作れない", h="途中を2〜3点刻みで取得し7.3条件まで外挿", g="計画確定時", t="平尾/木村", rd="", y="高"),
    # --- 設備 ---
    dict(b="設備", k="原料水を一気に絞れない（2段階の理由）", s="循環水の下限・設備制約", h="EG-1/EG-2の2段階で実施", g="計画確定時", t="平尾/木村", rd="", y="高"),
    dict(b="設備", k="コンデンサ・三重効用缶・熱交の負荷限界", s="M50×約1.05倍が限界の機器あり", h="OPERATION RANGEで成立性を確認", g="機器更新評価(8月末)", t="平尾/本田", rd="", y="中"),
    dict(b="設備", k="製品ポンプ吸入ストレーナ/NPSH", s="差圧上昇・閉塞の懸念", h="差圧を監視し基準超で中止判断", g="テスト期間中", t="平尾/製造", rd="", y="中"),
    dict(b="設備", k="出口戦略の新ライン（IER/吸着塔）の織込み先", s="サイズ大。起業に織込むか別途かの整理要", h="設置形態を整理", g="8/20ゲート", t="木村/基礎研", rd="", y="中"),
    # --- 会議 / SA ---
    dict(b="会議/SA", k="EG関係SAの起票（課1次・部1次・課3次の順）", s="EG関係はSA未実施で生産計画へ織込みづらい（EO-1は3次SA済）", h="SU前から逆算して起票・前倒し", g="SU前", t="木村/平尾", rd="", y="高"),
    dict(b="会議/SA", k="EO稼働50%の変更管理", s="モデルチェンジ大変更の一部", h="変更管理事項として整理", g="SA前", t="木村/平尾", rd="", y="高"),
    dict(b="会議/SA", k="インターロックタイマー20秒問題", s="設計思想が不明で単純な設定変更は不可", h="設計思想を解明（関係者ヒアリング）", g="3次SA前", t="平尾", rd="", y="中"),
    # --- 推進 / 調整 ---
    dict(b="推進/調整", k="テストの位置づけの一貫説明", s="必要性の見方に揺れ", h="「対策を効かせた状態で品質が持つか実機確認」で統一", g="7月上旬", t="平尾/木村", rd="", y="高"),
    dict(b="推進/調整", k="事業部（アルB）との個別議論", s="定例だけでは詰まらない", h="定例外で議論の場を設定", g="7月上旬", t="木村", rd="", y="高"),
    dict(b="推進/調整", k="M50(8/20ゲート)と税群DEG品質対応のスコープ分け", s="基礎研リソースが不足", h="スコープを分けて整理", g="8/20ゲート", t="木村/髙橋", rd="", y="中"),
    dict(b="推進/調整", k="基礎研の出口戦略リソース確保", s="人員減で逼迫", h="事業部・上位と調整", g="継続", t="木村", rd="", y="中"),
]

# ====== スケジュール（本筋のマイルストーン）======
SCHEDULE = [
    ("テストの位置づけ・コンセプトの統一", "7月上旬", "平尾/木村", "対策を効かせた状態で確認する建て付け"),
    ("事業部(アルB)との個別議論", "7月上旬", "木村", "定例外で論点を詰める"),
    ("触媒交換完了・EOスタート", "7/11見込み", "EO課/製造", "S/U手順は水を持ち込まない"),
    ("外販EOタンク繰りの成立", "7/5前後", "本田/遠藤/物流", "中間タンクのバッファ運用"),
    ("RD事前提出データの整理", "SU前/SA前", "基礎研/木村", "リン酸根拠・FALD/AALD根拠・IER挙動ほか"),
    ("色相悪化メカニズム仮説の擦り合わせ", "SU前", "基礎研/木村", "事実・仮説・対応で1枚に整理"),
    ("EG関係SAの起票（課1次・部1次・課3次の順）", "SU前", "木村/平尾", "EO-1は3次SA済、EG関係はこれから"),
    ("変更管理（EO稼働50%）・ILタイマー20秒の整理", "3次SA前", "平尾/木村", "設計思想の解明を含む"),
    ("カウンター3点の織込みとSOP反映", "7/27週", "平尾/木村/製造", "リン酸・塔頂圧ダウン・温度低下"),
    ("EG-1（低生産量）テスト", "7/28〜8/2", "平尾/木村/製造", "EO稼働50%"),
    ("EG-2（低水比）テスト", "8/3〜8/5", "平尾/木村/製造", "稼働70%近くへ"),
    ("DEGタンク組分け・東レ向け出荷不可期間の共有", "テスト期間中", "木村/遠藤/物流品証", "3期DEGタンク戻り時期の確認"),
    ("結果の評価（2〜3点刻み・7.3外挿）", "8/5以降", "平尾/木村", "幅を持たせて提示"),
    ("M50とDEG品質対応のスコープ分け・リソース確保", "8月", "木村/髙橋", "基礎研リソースの調整"),
    ("M50 8/20ゲート", "8/20", "平尾/木村/事業部", "ゲート審査"),
]


def write_sheet(ws, headers, widths, rows, note=None, group_col=None):
    ws.sheet_view.showGridLines = False
    r = 1
    if note:
        c = ws.cell(row=r, column=1, value=note)
        c.font = Font(name=FONT, size=9, color="595959")
        c.alignment = Alignment(horizontal="left", vertical="center")
        r += 1
    head_row = r
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=head_row, column=j, value=h)
        c.font = Font(name=FONT, size=10, bold=True)
        c.fill = fill(HEADER_FILL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    r = head_row + 1
    prev_group = None
    for row in rows:
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=r, column=j, value=val)
            c.font = Font(name=FONT, size=10)
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            c.border = border
        # 軽い区切り陰影（分類が変わる行）
        if group_col is not None:
            gv = row[group_col].split("/")[0]
            if gv != prev_group:
                for j in range(1, len(headers) + 1):
                    ws.cell(row=r, column=j).fill = fill(GROUP_FILL)
                prev_group = gv
        r += 1
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    ws.auto_filter.ref = "A%d:%s%d" % (head_row, get_column_letter(len(headers)), max(head_row, r - 1))


def main():
    wb = Workbook()

    # ① テスト概要
    ws0 = wb.active
    ws0.title = "テスト概要"
    write_sheet(ws0, ["項目", "内容"], [22, 92],
                [[k, v] for k, v in OVERVIEW],
                note="EG品質確認テスト(EG-1/EG-2) 本筋の整理（2026-06-30時点）")

    # ② 課題一覧
    ws1 = wb.create_sheet("課題一覧")
    h1 = ["No", "分類", "課題・論点", "制約", "対応方針", "期限", "担当", "RD", "優先度"]
    w1 = [4, 13, 30, 26, 30, 14, 13, 4, 7]
    rows1 = [[i, m["b"], m["k"], m["s"], m["h"], m["g"], m["t"], m["rd"], m["y"]]
             for i, m in enumerate(MASTER, start=1)]
    write_sheet(ws1, h1, w1, rows1, group_col=1,
                note="分類順（運転・品質・設備・会議・推進）。RD列○＝明日RDに事前提出を依頼。")

    # ③ RD事前提出依頼
    ws2 = wb.create_sheet("RD事前提出依頼")
    h2 = ["No", "RDに事前提出してほしい事項", "狙い・対応方針", "期限", "担当", "優先度"]
    w2 = [4, 38, 36, 14, 16, 7]
    rd = [m for m in MASTER if m["rd"] == "○"]
    rows2 = [[i, m["k"], m["h"], m["g"], m["t"], m["y"]] for i, m in enumerate(rd, start=1)]
    write_sheet(ws2, h2, w2, rows2,
                note="明日の品質定例で提示：RD(基礎研含む)に事前に出してほしいデータ・検証")

    # ④ スケジュール
    ws3 = wb.create_sheet("スケジュール")
    h3 = ["No", "マイルストーン", "時期", "担当", "補足"]
    w3 = [4, 38, 14, 18, 34]
    rows3 = [[i, m, t, o, n] for i, (m, t, o, n) in enumerate(SCHEDULE, start=1)]
    write_sheet(ws3, h3, w3, rows3,
                note="テスト工程とSA工程（課1次・部1次・課3次SAの順）のマイルストーン")

    out = os.path.join(OUT, "EG品質確認テスト_課題整理_20260630.xlsx")
    wb.save(out)
    print("saved:", out)
    print("概要:", len(OVERVIEW), "| 課題:", len(MASTER), "| RD:", len(rd), "| スケジュール:", len(SCHEDULE))


if __name__ == "__main__":
    main()
