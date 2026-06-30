# -*- coding: utf-8 -*-
"""EG品質確認テスト(EG-1/EG-2) 課題整理 Excel生成（1枚・つながり可視版）。

入力 : scratchpad/eg_design.json
       {zentai:[str], sections:[{title,nerai,rows:[{kadai,haikei,houshin,tunagari,kigen,tantou,rd,yusen}]}],
        local_todo:[str], zure_fix:[str]}
出力 : outputs/EG品質確認テスト_課題整理_20260630.xlsx （シートは1枚のみ）

方針 : 1枚に統合。上から「全体像(前提)→課題整理(テストの流れ順・つながり付)→ローカルで詰める事項(申し送り)」。
体裁 : モノクロ・マス結合なし・装飾/造語/矢印記号なし・出典/作成者表記なし。
"""
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "outputs")
SCRATCH = "/tmp/claude-0/-home-user-Base/78076d83-8708-5756-a6f1-f728259734b3/scratchpad"
os.makedirs(OUT, exist_ok=True)

FONT = "游ゴシック"
HEADER_FILL = "D9D9D9"   # 表ヘッダ
GROUP_FILL = "ECECEC"    # 区分見出し行
BLOCK_FILL = "F2F2F2"    # ブロック見出し
INK = "000000"
BORDER_CLR = "808080"
thin = Side(style="thin", color=BORDER_CLR)
border = Border(left=thin, right=thin, top=thin, bottom=thin)

HEAD = ["No", "課題・論点", "背景・制約", "対応方針", "つながり（何の前提／何に効く）",
        "期限", "担当", "RD", "優先"]
W = [4, 27, 25, 28, 26, 12, 13, 4, 6]
NC = len(HEAD)


def fill(c):
    return PatternFill("solid", fgColor=c)


def g(d, k):
    v = d.get(k, "")
    return "" if v is None else v


def est_height(values):
    lines = 1
    for val, w in zip(values, W):
        s = "" if val is None else str(val)
        cpl = max(4, int(w * 1.7))
        sub = sum(max(1, -(-len(p) // cpl)) for p in s.split("\n"))
        lines = max(lines, sub)
    return min(150, 15 * lines + 4)


def main():
    with open(os.path.join(SCRATCH, "eg_design.json"), encoding="utf-8") as f:
        D = json.load(f)

    wb = Workbook()
    ws = wb.active
    ws.title = "課題整理"
    ws.sheet_view.showGridLines = False
    r = 1

    def put(row, col, val, *, bold=False, size=10, fillc=None, bd=False,
            color=INK, wrap=True, va="top", ha="left"):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name=FONT, size=size, bold=bold, color=color)
        c.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
        if fillc:
            c.fill = fill(fillc)
        if bd:
            c.border = border
        return c

    # タイトル
    put(r, 1, "EG品質確認テスト(EG-1/EG-2) 課題整理　2026-06-30時点（ローカル継続用の申し送りを兼ねる）",
        bold=True, size=12, wrap=False)
    r += 2

    # 全体像（前提）
    put(r, 1, "■ 全体像（前提）", bold=True, size=11, fillc=BLOCK_FILL, wrap=False)
    r += 1
    for line in D.get("zentai", []):
        put(r, 2, "・" + line, wrap=False)
        ws.row_dimensions[r].height = 16
        r += 1
    r += 1

    # 課題整理（表ヘッダ）
    put(r, 1, "■ 課題整理（テストの流れ順）", bold=True, size=11, fillc=BLOCK_FILL, wrap=False)
    r += 1
    for j, h in enumerate(HEAD, start=1):
        put(r, j, h, bold=True, fillc=HEADER_FILL, bd=True, ha="center")
    r += 1

    no = 0
    for sec in D.get("sections", []):
        # 区分見出し行
        label = sec.get("title", "")
        if sec.get("nerai"):
            label = label + "　— " + sec["nerai"]
        for j in range(1, NC + 1):
            put(r, j, label if j == 1 else None, bold=True, fillc=GROUP_FILL, bd=True, wrap=False)
        r += 1
        # 課題行
        for row in sec.get("rows", []):
            no += 1
            vals = [no, g(row, "kadai"), g(row, "haikei"), g(row, "houshin"),
                    g(row, "tunagari"), g(row, "kigen"), g(row, "tantou"),
                    g(row, "rd"), g(row, "yusen")]
            for j, v in enumerate(vals, start=1):
                ha = "center" if j in (1, 8, 9) else "left"
                put(r, j, v, bd=True, ha=ha)
            ws.row_dimensions[r].height = est_height(vals)
            r += 1
    r += 1

    # ローカルで詰める事項（申し送り）
    put(r, 1, "■ ローカルで詰める事項（申し送り）", bold=True, size=11, fillc=BLOCK_FILL, wrap=False)
    r += 1
    for line in D.get("local_todo", []):
        put(r, 2, "・" + line, wrap=False)
        ws.row_dimensions[r].height = 16
        r += 1

    for j, w in enumerate(W, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    out = os.path.join(OUT, "EG品質確認テスト_課題整理_20260630.xlsx")
    wb.save(out)
    n_rows = sum(len(s.get("rows", [])) for s in D.get("sections", []))
    print("saved:", out)
    print("前提:", len(D.get("zentai", [])), "| 区分:", len(D.get("sections", [])),
          "| 課題:", n_rows, "| 申し送り:", len(D.get("local_todo", [])))


if __name__ == "__main__":
    main()
