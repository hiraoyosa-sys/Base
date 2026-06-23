# -*- coding: utf-8 -*-
"""
P-1507 技術評価 rev9 を構築する。
- rev8 の全シート/数式/キャッシュ値/書式/グラフを忠実に再現（xlsxwriterでキャッシュ値付き数式）
- 吐出圧・差圧にDCSタグが無い実態に合わせ「差圧直読」の結論を実運用(CV開度トレンド＋現場
  パトロール＋清掃周期管理)へ修正
- 新シート ⑤清掃周期：操油課250メッシュ実績を「ろ過面積比＋流量比」で当方周期へ換算する
  定量フレーム（★は操油課回答待ちのプレースホルダ）を追加
"""
import openpyxl, xlsxwriter

SRC = "outputs/P1507ストレーナ変更_技術評価_rev10.xlsx"
DST = "outputs/P1507ストレーナ変更_技術評価_rev11.xlsx"

wbf = openpyxl.load_workbook(SRC, data_only=False)
wbv = openpyxl.load_workbook(SRC, data_only=True)

# 「差圧直読」が不可能になった実態に合わせた結論差し替え
REPLACE = {
 ("①NPSH", "A11"):
   "結論：通常運転はクリーン3.6kPa＝余力の約3%。約83%閉塞で余裕喪失。"
   "差圧の直読タグは無いため、CV開度トレンド＋現場パトロール＋清掃周期管理(⑤)で監視。",
 ("②ストレーナ", "A46"):
   "結論：ろ過面積≈101cm²。初期差圧は小。約83%閉塞で余裕喪失。"
   "監視はCV開度トレンド＋現場パトロール(1日1回)＋清掃周期見積り(⑤)。",
 ("④物性(Aspen)", "A44"):
   "→ 対策：~50℃以上へ暖機/低流量起動/冷態は40メッシュ。起動時は閉塞速度を現場で確認。",
 ("③CV監視(両弁)", "A29"):
   "→ +6〜8%は通常の流量変動(±5〜10%)と同程度。差圧直読タグは無いため、主管理は"
   "清掃周期(⑤)＋現場パトロール、CV開度トレンドは“弱い兆候”として補助に使う。",
 ("①NPSH", "A2"):
   "考え方：塔頂圧≒蒸気圧で相殺→NPSHa≈静水頭で一定。削るのはストレーナΔPのみ。",
 ("0_読み方", "B3"):
   "黄=入力(出典)。白の結果はクリックで数式。物性=Aspen係数表参照。"
   "CV=AGVB実流量特性カーブ補間＋DCS実機照合(③)。差圧直読タグが無いため監視は清掃周期(⑤)主体。",
}

def style_to_fmt(book, cell, cache):
    """openpyxl セル書式 → xlsxwriter Format（キャッシュして再利用）"""
    f = cell.font; fl = cell.fill; al = cell.alignment; bd = cell.border
    d = {}
    if f.bold: d["bold"] = True
    if f.size and f.size != 11: d["font_size"] = f.size
    # 文字は黒ベース：青/白などのフォント色は捨てる（既定の黒）
    if fl is not None and fl.patternType:
        fg = fl.fgColor
        rgb = getattr(fg, "rgb", None)
        if isinstance(rgb, str):
            hexv = rgb[-6:].upper()
            if hexv == "FFF2CC":              # 入力＝黄色は維持
                d["bg_color"] = "#FFF2CC"; d["border"] = 1
            elif hexv == "1F4E79":            # 旧・濃紺ヘッダ → 白地・黒太字・罫線
                d["bold"] = True; d["border"] = 1
            # それ以外の塗りは白地に
    # 元の罫線は薄罫線として維持（表のグリッド感）
    if any(s and s.style for s in (bd.left, bd.right, bd.top, bd.bottom)):
        d["border"] = 1
    if al.wrap_text: d["text_wrap"] = True
    if al.horizontal: d["align"] = al.horizontal
    if al.vertical: d["valign"] = al.vertical
    key = tuple(sorted(d.items()))
    if key not in cache:
        cache[key] = book.add_format(d) if d else book.add_format()
    return cache[key]

def copy_sheet(book, name, fmt_cache):
    wsf = wbf[name]; wsv = wbv[name]
    ws = book.add_worksheet(name)
    # 列幅
    for letter, dim in wsf.column_dimensions.items():
        if dim.width:
            ci = openpyxl.utils.column_index_from_string(letter) - 1
            ws.set_column(ci, ci, dim.width)
    for row in wsf.iter_rows():
        for c in row:
            if c.value is None:
                continue
            r, col = c.row - 1, c.column - 1
            fmt = style_to_fmt(book, c, fmt_cache)
            val = c.value
            # 結論等の差し替え
            if (name, c.coordinate) in REPLACE:
                val = REPLACE[(name, c.coordinate)]
                ws.write(r, col, val, fmt); continue
            if isinstance(val, str) and val.startswith("="):
                cached = wsv[c.coordinate].value
                ws.write_formula(r, col, val, fmt, cached if cached is not None else 0)
            else:
                ws.write(r, col, val, fmt)
    return ws

book = xlsxwriter.Workbook(DST, {"strings_to_numbers": False})
fmt_cache = {}

# ---- 既存5シートを順に再現 ----
ws_read   = copy_sheet(book, "0_読み方", fmt_cache)
ws_npsh   = copy_sheet(book, "①NPSH", fmt_cache)
ws_str    = copy_sheet(book, "②ストレーナ", fmt_cache)
ws_cv     = copy_sheet(book, "③CV監視(両弁)", fmt_cache)
ws_prop   = copy_sheet(book, "④物性(Aspen)", fmt_cache)

# 0_読み方 に ⑤ の行を追記
ws_read.write(9, 1, "⑤清掃周期")
ws_read.write(9, 2, "操油課250メッシュの差圧上昇実績→ろ過面積比で当方の清掃周期を換算(★回答待ち)")

# ============== 新シート ⑤清掃周期（黒文字・白地・入力のみ黄色） ==============
F = lambda **k: book.add_format(k)
f_title  = F(bold=True, font_size=14)
f_note   = F(text_wrap=True, valign="top")
f_hdr    = F(bold=True, text_wrap=True, border=1, align="center")   # 白地・黒太字・罫線
f_sec    = F(bold=True)
f_in     = F(bg_color="#FFF2CC", border=1)             # 入力(黄)
f_instar = F(bg_color="#FFF2CC", border=1)             # ★も黄色のみ(文字は黒)
f_out    = F(border=1)                                 # 計算結果(白)
f_outb   = F(border=1, bold=True)
f_txt    = F(border=1)
f_concl  = F(bold=True, text_wrap=True, valign="top")

def A1(rr, col=1): return f"{chr(65+col)}{rr+1}"   # 0-index行 → Excel A1参照(既定B列)

w = book.add_worksheet("⑤清掃周期")
w.set_column(0, 0, 30); w.set_column(1, 1, 13); w.set_column(2, 2, 8); w.set_column(3, 3, 54)

w.write(0, 0, "⑤ 清掃周期の見積り（操油課250メッシュの差圧上昇実績 → ろ過面積比換算）", f_title)
w.merge_range(1, 0, 1, 3,
    "考え方：清掃周期は『差圧がどの期間でどれくらい上がるか（上昇速度）』で決まる。"
    "操油課250メッシュの差圧上昇速度を、ろ過面積比で当方へ換算する（面積が大きいほど"
    "上昇は緩やか＝長持ち）。当方の許容差圧上昇はキャビ限界117−クリーン3.6＝約113kPa。"
    "★＝操油課回答待ち(木村課長へ照会済)。初期は安全側で早めに点検し、起動時の差圧の"
    "立ち方とCV開度トレンドで実機補正する。", f_note)
w.set_row(1, 70)

r = 3
w.write(r, 0, "■ 入力：操油課 250メッシュ 差圧上昇実績（★は回答待ち）", f_sec); r += 1
for j, h in enumerate(["項目", "値", "単位", "出典・注記"]):
    w.write(r, j, h, f_hdr)
r += 1
R_dPobs, R_tobs, R_Aref = r, r+1, r+2
rows_in = [
 ("操油課 差圧上昇量 ΔP_obs",  50, "kPa",  "★操油課回答待ち（清掃判断までに上がる差圧。暫定50）"),
 ("操油課 計測期間 t_obs",     60, "日",   "★操油課回答待ち（その差圧が上がるのに要した期間。暫定60）"),
 ("操油課 ろ過面積 A_ref",     50, "cm²",  "★操油課回答待ち（暫定50）"),
]
for i, (a, b, c, d) in enumerate(rows_in):
    w.write(r+i, 0, a, f_txt); w.write(r+i, 1, b, f_instar)
    w.write(r+i, 2, c, f_txt); w.write(r+i, 3, d, f_txt)
r += len(rows_in)
cdP, ctobs, cAref = A1(R_dPobs), A1(R_tobs), A1(R_Aref)
R_rateRef = r
w.write(r, 0, "操油課 差圧上昇速度", f_txt)
w.write_formula(r, 1, f"={cdP}/{ctobs}", f_out, round(50/60, 3))
w.write(r, 2, "kPa/日", f_txt); w.write(r, 3, "=ΔP_obs/t_obs", f_txt)
r += 2

w.write(r, 0, "■ 当方条件（本検討より自動引用）", f_sec); r += 1
for j, h in enumerate(["項目", "値/式", "単位", "出典・式"]):
    w.write(r, j, h, f_hdr)
r += 1
R_Aus = r
w.write(r, 0, "当方 ろ過面積 A_us", f_txt)
w.write_formula(r, 1, "='②ストレーナ'!B18", f_out, 101.3)
w.write(r, 2, "cm²", f_txt); w.write(r, 3, "②ストレーナ ろ過面積(円錐式)", f_txt)
r += 1
R_budget = r
w.write(r, 0, "当方 許容差圧上昇 ΔP_lim", f_txt)
w.write_formula(r, 1, "='①NPSH'!B9-'②ストレーナ'!B10", f_out, round(116.9-3.56, 1))
w.write(r, 2, "kPa", f_txt); w.write(r, 3, "=キャビ限界(①B9) − クリーンΔP(②B10)", f_txt)
r += 2
cAus, cBudget, cRateRef = A1(R_Aus), A1(R_budget), A1(R_rateRef)

w.write(r, 0, "■ ろ過面積比換算による推定清掃周期", f_sec); r += 1
for j, h in enumerate(["項目", "値/式", "単位", "式・注記"]):
    w.write(r, j, h, f_hdr)
r += 1
R_ratioA = r
w.write(r, 0, "面積比 (操油課/当方)", f_txt)
w.write_formula(r, 1, f"={cAref}/{cAus}", f_out, round(50/101.3, 4))
w.write(r, 2, "-", f_txt); w.write(r, 3, "=A_ref/A_us（上昇速度の換算係数）", f_txt)
r += 1
cRatioA = A1(R_ratioA)
R_rateUs = r
w.write(r, 0, "当方 差圧上昇速度", f_txt)
w.write_formula(r, 1, f"={cRateRef}*{cRatioA}", f_out, round((50/60)*(50/101.3), 4))
w.write(r, 2, "kPa/日", f_txt); w.write(r, 3, "=操油課速度×面積比（面積大→緩やか）", f_txt)
r += 1
cRateUs = A1(R_rateUs)
R_period = r
w.write(r, 0, "推定清掃周期", f_txt)
w.write_formula(r, 1, f"={cBudget}/{cRateUs}", f_outb, round((116.9-3.56)/((50/60)*(50/101.3)), 1))
w.write(r, 2, "日", f_txt); w.write(r, 3, "=許容差圧上昇 / 当方上昇速度", f_txt)
r += 1
cPeriod = A1(R_period)
R_first = r
w.write(r, 0, "初回点検時期(安全側)", f_txt)
w.write_formula(r, 1, f"={cPeriod}*0.5", f_outb, round((116.9-3.56)/((50/60)*(50/101.3))*0.5, 1))
w.write(r, 2, "日", f_txt); w.write(r, 3, "初期は推定周期の1/2で点検→実績で延伸", f_txt)
r += 2

w.merge_range(r, 0, r, 3,
    "運用：①操油課の差圧上昇実績を入力 → ②面積比で当方の上昇速度・推定周期を算出 → "
    "③初期は推定周期の1/2で開放点検し差圧の立ち方を実測 → ④起動時の差圧上昇とCV開度トレンド"
    "(DCS)・現場パトロール(1日1回)で周期を実機補正 → ⑤以後は補正済み周期で定期清掃。"
    "差圧の直読タグが無いため、定量管理はこの清掃周期が主軸。"
    "（注）固形分濃度・流量が操油課と大きく異なる場合は起動時実測で補正する。", f_concl)
w.set_row(r, 70)

# ============== 既存グラフ4枚を再現 ==============
# chart1: ② 閉塞率→差圧
c1 = book.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
c1.add_series({"name": "ストレーナ差圧",
    "categories": ["②ストレーナ", 34, 0, 42, 0], "values": ["②ストレーナ", 34, 2, 42, 2]})
c1.add_series({"name": "NPSH余力(限界117)",
    "categories": ["②ストレーナ", 34, 0, 42, 0], "values": ["②ストレーナ", 34, 3, 42, 3]})
c1.set_title({"name": "閉塞率 → ストレーナ差圧（約83%で余裕喪失）"})
c1.set_x_axis({"name": "閉塞率 b [%]"})
c1.set_y_axis({"name": "差圧 [kPa]"})
ws_str.insert_chart("F12", c1, {"x_scale": 1.2, "y_scale": 1.2})

# chart2: ③ AGVB等％カーブ
c2 = book.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
c2.add_series({"name": "AGVB 等％",
    "categories": ["③CV監視(両弁)", 4, 6, 13, 6], "values": ["③CV監視(両弁)", 4, 7, 13, 7]})
c2.set_title({"name": "AGVB 等％ 流量特性（開度 vs Cv%）"})
c2.set_x_axis({"name": "開度 [%]"})
c2.set_y_axis({"name": "Cv [%]"})
c2.set_legend({"none": True})
ws_cv.insert_chart("G16", c2, {"x_scale": 1.2, "y_scale": 1.2})

# chart3: ④ 粘度の温度依存（対数）
c3 = book.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
c3.add_series({"name": "μ(T)",
    "categories": ["④物性(Aspen)", 25, 0, 35, 0], "values": ["④物性(Aspen)", 25, 1, 35, 1]})
c3.set_title({"name": "DEG 粘度の温度依存（Aspen MULDIP式）"})
c3.set_x_axis({"name": "温度 [℃]"})
c3.set_y_axis({"name": "粘度 μ [cP]（対数）", "log_base": 10})
c3.set_legend({"none": True})
ws_prop.insert_chart("I4", c3, {"x_scale": 1.15, "y_scale": 1.15})

# chart4: ④ 低温ΔP(T) と 余力
c4 = book.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
c4.add_series({"name": "250ΔP(T)",
    "categories": ["④物性(Aspen)", 25, 0, 35, 0], "values": ["④物性(Aspen)", 25, 4, 35, 4]})
c4.add_series({"name": "限界117",
    "categories": ["④物性(Aspen)", 25, 0, 35, 0], "values": ["④物性(Aspen)", 25, 6, 35, 6]})
c4.set_title({"name": "低温での250ストレーナΔP(T) と NPSH余力"})
c4.set_x_axis({"name": "温度 [℃]"})
c4.set_y_axis({"name": "差圧/余力 [kPa]"})
ws_prop.insert_chart("I20", c4, {"x_scale": 1.15, "y_scale": 1.15})

book.close()
print("wrote", DST)
