#!/usr/bin/env python3
# EG低稼働品質確認テスト SA資料への差込スライド案（LERA運転モード整理ほか5枚）
# ベース: GDrive スライドテンプレート/MCG_jp_rev20260628.pptx（remote-slide-skill手順準拠）
import copy
import os
import re
import shutil
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN

SCRATCH = "/tmp/claude-0/-home-user-Base/201f7e38-9b57-5ef9-a463-64d4d9a9f041/scratchpad"
TEMPLATE = os.path.join(SCRATCH, "template_src.pptx")
OUT_DIR = "/home/user/Base/outputs"
OUT_PPTX = os.path.join(OUT_DIR, "20260715_EG低稼働テストSA_LERA整理_差込案.pptx")

FONT = "BIZ UDPゴシック"
RED = RGBColor(0xFF, 0x00, 0x00)
BLUE = RGBColor(0x00, 0x00, 0xFF)
GRAY = RGBColor(0x59, 0x59, 0x59)
BLACK = RGBColor(0x00, 0x00, 0x00)
HEAD_FILL = "F2F2F2"


def style_run(run, size=16, bold=False, color=None):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.name = FONT
    if color is not None:
        f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        if latin is not None:
            latin.addnext(ea)
        else:
            rPr.append(ea)
    ea.set("typeface", FONT)


def add_text(slide, x, y, w, h, items, space_after=4):
    """items: list of paragraphs; each paragraph = list of (text, kw) runs"""
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for para in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(space_after)
        for text, kw in para:
            r = p.add_run()
            r.text = text
            style_run(r, **kw)
    return tb


def P(text, **kw):
    return [(text, kw)]


def make_table(slide, x, y, w, rows_spec, col_w, row_h, header=True):
    n_rows = len(rows_spec)
    n_cols = len(col_w)
    gf = slide.shapes.add_table(n_rows, n_cols, Cm(x), Cm(y), Cm(w), Cm(row_h * n_rows))
    table = gf.table
    tbl = gf._element.graphic.graphicData.tbl
    tblPr = tbl.find(qn("a:tblPr"))
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    style_id = tblPr.find(qn("a:tableStyleId"))
    if style_id is not None:
        style_id.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # No Style, Table Grid
    for i, wcm in enumerate(col_w):
        table.columns[i].width = Cm(wcm)
    for i in range(n_rows):
        table.rows[i].height = Cm(row_h)
    for ri, row in enumerate(rows_spec):
        for ci, cell_spec in enumerate(row):
            cell = table.cell(ri, ci)
            cell.margin_left = Cm(0.15)
            cell.margin_right = Cm(0.15)
            cell.margin_top = Cm(0.05)
            cell.margin_bottom = Cm(0.05)
            if header and ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(HEAD_FILL)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            tf = cell.text_frame
            tf.word_wrap = True
            if isinstance(cell_spec, str):
                cell_spec = [[(cell_spec, {})]]
            first = True
            for para in cell_spec:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                for text, kw in para:
                    r = p.add_run()
                    r.text = text
                    kw = dict(kw)
                    kw.setdefault("size", 16)
                    if header and ri == 0:
                        kw.setdefault("bold", True)
                    style_run(r, **kw)
    return gf


def message_line(slide, text):
    add_text(slide, 0.7, 1.95, 32.4, 1.0, [P(text, size=16, bold=True)])


def note_line(slide, text, y=17.15):
    add_text(slide, 0.7, y, 32.4, 0.9, [P(text, size=12, color=GRAY)])


def main():
    prs = Presentation(TEMPLATE)
    slides = prs.slides
    base = slides[1]  # 1_basic_text: title / footer / page no / object ph

    # --- 雛形を複製して計5枚にする（add_slideで枠を確保→中身をbaseから深いコピー） ---
    def dup(src):
        new = slides.add_slide(src.slide_layout)
        for sh in list(new.shapes):
            sh._element.getparent().remove(sh._element)
        for sh in src.shapes:
            new.shapes._spTree.append(copy.deepcopy(sh._element))
        return new

    pages = [base] + [dup(base) for _ in range(4)]

    # 表紙（テンプレ1枚目）を削除
    sldIdLst = slides._sldIdLst
    first = list(sldIdLst)[0]
    rId = first.get(qn("r:id"))
    prs.part.drop_rel(rId)
    sldIdLst.remove(first)

    # 各ページ: 空のOBJECTプレースホルダを除去
    for s in pages:
        for sh in list(s.shapes):
            if sh.is_placeholder and sh.placeholder_format.idx == 10:
                sh._element.getparent().remove(sh._element)

    # ============ 1. テストの目的（修正文言案） ============
    s = pages[0]
    s.shapes.title.text = "テストの目的"
    message_line(s, "対策を織り込んだM50相当運転で、製品品質が規格・運転管理値に収まり、想定外の変化がないことを実機で確認する")
    add_text(s, 1.2, 3.6, 31.0, 5.0, [
        P("背景", size=18, bold=True),
        P("・M50運転では系内の不純物濃度が上がり、製品品質の確保が課題となる", size=18),
        P("・対策（リン酸塩添加・MEG塔温度低減）を検討しており、有効と考えている", size=18),
    ], space_after=8)
    add_text(s, 1.2, 8.6, 31.0, 7.0, [
        P("実機テスト目的", size=18, bold=True),
        P("・M50相当の運転条件（低稼働・低水比）において、対策を織り込んだ運転で以下を確認する", size=18),
        P("　確認①：製品（MEG・DEG・TEG）の品質が規格・運転管理値に収まること", size=18),
        P("　確認②：系内の不純物バランスが想定どおりであること（想定外の変化がないことの確認）", size=18),
    ], space_after=8)

    # ============ 2. LERA運転モードの比較 ============
    s = pages[1]
    s.shapes.title.text = "LERA運転モードの比較"
    message_line(s, "どちらのモードでもテスト目的は達成できる。現場負荷とリスクの見合いで選択したい（相談）")
    rows = [
        ["観点", "LERA to EG（現計画）", "LERA to EO（切替案）"],
        ["テスト成立性",
         "成立（EG行きEO＝精製塔2.3＋LERA 5.2 t/h）",
         "成立（EG行きEO＝精製塔7.5 t/h）"],
        ["制御・安定運転",
         [[("課題：", {"bold": True, "color": RED}), ("EO 2.3t/h＝CV開度2〜3%・LL近接。成立性を定量評価中（次頁）", {})]],
         [[("CV通常開度域＝制御課題は解消", {"color": BLUE})]]],
        ["保安（I/L）",
         "FK1432 LL 6.0→0.6 t/hの変更要（期間限定・終了後復旧）",
         [[("I/L変更不要", {"color": BLUE})]]],
        ["非定常操作",
         "切替操作なし",
         [[("課題：", {"bold": True, "color": RED}), ("切替操作あり（半日・ボード4名専念）＝役割を事前指名", {})]]],
        ["現場負荷",
         "通常直で対応可",
         "切替日に応援・時間外が必要（人数・日程は運転と調整中）"],
        ["生産・外販",
         "外販253t/d・50%稼働を維持",
         "切替直後に外販約10t/hの一時ダウン（外販バランスへ織込み）。EG2は70%稼働可"],
        ["経済性",
         "増分なし",
         "蒸気＋10t/h＝＋0.6〜1.1百万円/日（約3週間で12〜17百万円）"],
        ["開始時期",
         "7/28開始可（現計画）",
         "制御性評価と切替準備の完了後＝1週ずらし案とセット"],
    ]
    make_table(s, 0.7, 3.05, 32.4, rows, col_w=[4.6, 13.9, 13.9], row_h=1.55)
    note_line(s, "決め手＝EG行きEO 2.3t/hの制御性評価（次頁）。切替の場合は実施時期と合わせて判断する")

    # ============ 3. EG行きEO 2.3t/hの制御性評価 ============
    s = pages[2]
    s.shapes.title.text = "EG行きEO 2.3t/hの制御性評価（LERA to EG継続の成立条件）"
    message_line(s, "3段の問いで定量評価し、その結果でLERA運転モードを確定する")
    add_text(s, 1.2, 3.4, 31.0, 11.5, [
        P("① トータルEO比率制御（カスケード）を切れるか", size=18, bold=True),
        P("　・LERA吸収量の変動が乗るため制御幅が大きい＝切離しの可否をトレンドで確認（※トレンド貼付）", size=16),
        P("② 単独FC（FC1402）で2.3t/hを制御できるか", size=18, bold=True),
        P("　・判断基準：FK1432指示ブレ（実測±0.2〜0.4t/h）に対する運転点2.3t/h・LL 0.6t/hの余裕", size=16),
        P("　・CV開度2〜3%域での制御実績の有無", size=16),
        P("③ （②が不可の場合）マニュアル一定でテストが成立するか", size=18, bold=True),
        P("　・成立の最低条件＝流量指示が正常範囲で見えていること", size=16),
    ], space_after=10)
    add_text(s, 1.2, 13.0, 31.0, 2.2, [
        [("結論の出し方：", {"size": 16, "bold": True}),
         ("評価完了の期日を決めて運転と共有。不成立ならLERA to EO切替（前頁）と実施時期見直しをセットで判断", {"size": 16})],
    ])
    note_line(s, "※現状は定性評価にとどまるため、指示ブレと余裕の定量比較で言い切れる形にする")

    # ============ 4. 実施時期の決め方 ============
    s = pages[3]
    s.shapes.title.text = "実施時期の決め方"
    message_line(s, "時期はオフタンク余力・現場人繰り・DEGタンク繰りの3要因で決まる。8月頭開始なら3要因とも余裕を持って成立の見込み（調整中）")
    rows = [
        ["決定要因", "7/28開始（現計画）", "8月頭開始（1週ずらし）"],
        ["① オフタンク余力（150〜200t）",
         "7/21回収開始・約6日間＝到達ぎりぎり",
         [[("余力増", {"color": BLUE})]]],
        ["② 現場の人繰り・残業",
         [[("課題：", {"bold": True, "color": RED}), ("SU直後で時間外がかさむ（事前教育は時間外で7/23〜24調整中）", {})]],
         [[("緩和", {"color": BLUE})]]],
        ["③ DEGタンク繰り・盆中の出荷",
         "盆前後の分割実施が前提（T-604・保管評価8日）",
         [[("SPタンク移送・スポット船の確度が上がる", {"color": BLUE})]]],
    ]
    make_table(s, 0.7, 3.6, 32.4, rows, col_w=[8.0, 12.2, 12.2], row_h=2.0)
    add_text(s, 1.2, 12.6, 31.0, 4.5, [
        P("合わせて整合を取る事項", size=16, bold=True),
        P("・抜出比率0.82の検証（1日で実施可）＝LERA to EG期間中に実施する", size=16),
        P("・M100テストとは干渉しない（7/27まで周期制約のため大きな条件変更なし）", size=16),
        P("・後ろ倒しにするほど検証後の対策検討期間が短くなる（8月中の実施が限度）", size=16),
    ], space_after=6)

    # ============ 5. 資料構成の見直し案（参考） ============
    s = pages[4]
    s.shapes.title.text = "資料構成の見直し案（参考）"
    message_line(s, "目的で立てたキーワード（規格・運転管理値／不純物バランス／想定内外）を後段で回収する並びに組み替える")
    add_text(s, 1.2, 3.3, 31.0, 14.0, [
        P("１．目的（P1改）＝キーワードの布石：規格・運転管理値／不純物バランス／想定内外", size=16),
        P("２．課題（P2〜4：M50課題・足元の不純物・GALD影響予測）", size=16),
        P("３．対策と有効性（P5〜7：リン酸塩・MEG塔圧ダウン・織込み事項）", size=16),
        P("４．テストコンセプト・条件（P8〜10：運転マップで「実績データが少ない＝想定外がないことの確認」を明記）", size=16),
        P("５．実施方式の選択＝LERA（比較 → 制御性評価 → 切替作業 → 変動費）", size=16),
        P("６．変更点に対する対応（P27〜42：運転・保安・品質）", size=16),
        P("７．テスト計画（実施時期 → 計画 → 監視項目 → 分析 → 関係先連絡P80）", size=16),
        P("８．SA総合評価（P54〜56）", size=16),
        P("補足：APHAメカニズム（P15）・旧条件図（P76〜79）は付録へ", size=16),
    ], space_after=9)

    # ============ メタ情報（AI痕跡除去） ============
    cp = prs.core_properties
    cp.author = "平尾　一陽"
    cp.last_modified_by = "平尾　一陽"
    cp.title = "EG低稼働品質確認テスト 差込スライド案"
    cp.comments = ""
    cp.category = ""

    os.makedirs(OUT_DIR, exist_ok=True)
    tmp = OUT_PPTX + ".tmp"
    prs.save(tmp)

    # webextensions（アドイン痕跡）を除去して再パッケージ
    src = zipfile.ZipFile(tmp)
    with zipfile.ZipFile(OUT_PPTX, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename.startswith("ppt/webextensions/"):
                continue
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                root = etree.fromstring(data)
                for el in list(root):
                    pn = el.get("PartName") or ""
                    ct = el.get("ContentType") or ""
                    if "webextension" in pn or "webextension" in ct or "taskpanes" in pn:
                        root.remove(el)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "ppt/_rels/presentation.xml.rels":
                root = etree.fromstring(data)
                for el in list(root):
                    if "webextensions" in (el.get("Target") or ""):
                        root.remove(el)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            dst.writestr(item, data)
    src.close()
    os.remove(tmp)
    print("saved:", OUT_PPTX, os.path.getsize(OUT_PPTX), "bytes")


if __name__ == "__main__":
    main()
