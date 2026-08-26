#!/usr/bin/env python3
"""分割スキャンされた長尺 P&ID を 1 枚に合成する。

対象: B240-01200 / P&ID No.120-6001
      UNIT 1600 EG BLEED RECOVERY (MYKK PROCESS) UNIT
      3 ページ (A4 縦 1 + A3 横 2、いずれも 200dpi) に分割スキャンされたもの。

手順
  1. 各ページから埋め込み JPEG を取り出す(再圧縮しない)。
  2. 各シートの図枠上下の罫線を実測し、上罫線 → Y_TOP、下罫線 → Y_BOT に
     なるよう列ごとに縦方向を補正する(傾き・台形ひずみ・紙の波打ちを同時に除去)。
  3. 継ぎ目は図面上で切れている線番号文字から求めた実測オフセットで突き合わせる。
        seam1  3/4"-CWS-6003-0A|1A     p1 x=1537  <->  p2 x=397
        seam2  5"-|CWS-6001-0A1A       p2 x=3147  <->  p3 x=106
  4. 図枠の外側 (スキャン余白・パンチ穴・ページ端の影) を落として 1 枚に切り出し、
     200dpi 相当の単ページ PDF として書き出す。
"""
import json
import cv2
import numpy as np
import pymupdf

SRC_PDF = "/root/.claude/uploads/6b876591-8bbc-5e5d-b3e5-ab1bceebfbfc/df84ebd3-200___.pdf"
OUT_DIR = "/home/user/Base/outputs"
DPI = 200.0

# ページ内の埋め込み画像 xref (左 → 右)
SHEETS = [("p1", 0, 8), ("p2", 1, 14), ("p3", 2, 17)]

# 各シートを合成キャンバスに置く x オフセット (実測)
DX = {"p1": 0, "p2": 1140, "p3": 1140 + 3041}
# 各シートの使用範囲 (元画像 x)。スキャン余白・ページ端の影を除く
CLIP = {"p1": (0, 1537), "p2": (380, 3147), "p3": (100, 3310)}
# 合成時の優先順 (後に描いたものが上)
ORDER = ["p3", "p2", "p1"]

# 図枠罫線の実測に使う探索範囲
TOPX = {"p1": (150, 1520), "p2": (400, 3145), "p3": (110, 3160)}
BOTX = {"p1": (150, 1520), "p2": (400, 3145), "p3": (120, 2620)}   # p3 は表題欄を避ける
GUESS = {"p1": (139, 2270), "p2": (136, 2268), "p3": (115, 2256)}

Y_TOP, FRAME_H = 60.0, 2136.0          # 出力キャンバス上の図枠位置と高さ
MARGIN = 26                            # 図枠の外側に残す余白 (px)


def extract_sheets():
    doc = pymupdf.open(SRC_PDF)
    out = {}
    for name, page, xref in SHEETS:
        info = doc.extract_image(xref)
        buf = np.frombuffer(info["image"], np.uint8)
        out[name] = cv2.imdecode(buf, cv2.IMREAD_COLOR)
    return out


def white_balance(img):
    """紙の地色をシート間でそろえる(チャンネル別。線の濃さは変えない)。"""
    out = img.astype(np.float32)
    for ch in range(3):
        paper = max(float(np.percentile(out[:, :, ch], 85)), 1.0)
        out[:, :, ch] *= 252.0 / paper
    return np.clip(out, 0, 255).astype(np.uint8)


def trace_rule(gray, y_start, x0, x1, step=50, wid=80, win=30, min_frac=0.8):
    """横方向に長い罫線を追跡して (x, y) を返す。"""
    pts, y = [], y_start
    for x in range(x0, x1 - wid, step):
        band = (gray[max(0, y - win):y + win, x:x + wid] < 170).astype(np.uint8)
        prof = band.sum(axis=1) / wid
        if prof.max() < min_frac:
            continue
        i = int(np.argmax(prof))
        lo, hi = max(0, i - 4), i + 5
        seg = prof[lo:hi]
        yy = float((np.arange(lo, hi) * seg).sum() / seg.sum()) + max(0, y - win)
        pts.append((x + wid / 2.0, yy))
        y = int(round(yy))
    return np.array(pts)


def robust_curve(pts, deg=2, thr=4.0, smooth=5):
    """外れ値を落として移動平均でならした罫線カーブを返す。"""
    x, y = pts[:, 0], pts[:, 1]
    coef = np.polyfit(x, y, deg)
    for _ in range(4):
        res = y - np.polyval(coef, x)
        keep = np.abs(res) < max(thr, 3 * res.std())
        if keep.sum() < deg + 4 or keep.all():
            break
        x, y = x[keep], y[keep]
        coef = np.polyfit(x, y, deg)
    ys = np.convolve(y, np.ones(smooth) / smooth, mode="same")
    h = smooth // 2
    ys[:h], ys[-h:] = y[:h], y[-h:]
    return x, ys


def curve_at(xs, cx, cy):
    """実測範囲の外は端の傾きで直線外挿する。"""
    out = np.interp(xs, cx, cy)
    for lo, sl in ((True, slice(0, 4)), (False, slice(-4, None))):
        a, b = np.polyfit(cx[sl], cy[sl], 1)
        m = xs < cx[0] if lo else xs > cx[-1]
        out[m] = a * xs[m] + b
    return out


def main():
    sheets = {k: white_balance(v) for k, v in extract_sheets().items()}
    rules = {}
    for k, img in sheets.items():
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        top = robust_curve(trace_rule(gray, GUESS[k][0], *TOPX[k]))
        bot = robust_curve(trace_rule(gray, GUESS[k][1], *BOTX[k]))
        rules[k] = (top, bot)
        print(f"{k}: 図枠 上罫線 {top[1].mean():6.1f}  下罫線 {bot[1].mean():7.1f} "
              f"(枠高 {bot[1].mean() - top[1].mean():.1f}px)")

    W = max(DX[k] + CLIP[k][1] for k in sheets) + 10
    H = int(Y_TOP + FRAME_H + Y_TOP)
    canvas = np.full((H, W, 3), 255, np.uint8)
    ys = np.arange(H, dtype=np.float32)

    for k in ORDER:
        img = sheets[k]
        (tx, ty), (bx, by) = rules[k]
        x0, x1 = CLIP[k]
        xs_out = np.arange(DX[k] + x0, DX[k] + x1, dtype=np.float32)
        xs_src = xs_out - DX[k]
        y_t = curve_at(xs_src, tx, ty)
        y_b = curve_at(xs_src, bx, by)
        # 出力 Y_TOP..Y_TOP+FRAME_H を、その列の 上罫線..下罫線 に対応させる
        scale = (y_b - y_t) / FRAME_H
        map_y = (y_t[None, :] + (ys[:, None] - Y_TOP) * scale[None, :]).astype(np.float32)
        map_x = np.repeat(xs_src[None, :], H, axis=0).astype(np.float32)
        patch = cv2.remap(img, map_x, map_y, cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))
        canvas[:, int(xs_out[0]):int(xs_out[0]) + patch.shape[1]] = patch
        print(f"{k}: 合成 x {int(xs_out[0])}..{int(xs_out[-1])}")

    # ---- 図枠の左右罫線を検出して切り出す ----
    ink = (cv2.cvtColor(canvas, cv2.COLOR_BGR2GRAY) < 170).astype(np.uint8)
    col = ink[int(Y_TOP) + 10:int(Y_TOP + FRAME_H) - 10, :].mean(axis=0)
    cand = np.where(col > 0.5)[0]
    left, right = int(cand.min()), int(cand.max())
    print(f"図枠: 左罫線 x={left}  右罫線 x={right}  "
          f"({right - left} x {int(FRAME_H)} px)")
    x0, x1 = max(0, left - MARGIN), min(W, right + MARGIN + 1)
    y0, y1 = int(Y_TOP) - MARGIN, int(Y_TOP + FRAME_H) + MARGIN + 1
    sheet = canvas[y0:y1, x0:x1]

    png = f"{OUT_DIR}/B240-01200_PID_merged.png"
    cv2.imwrite(png, sheet)

    h, w = sheet.shape[:2]
    ok, jpg = cv2.imencode(".jpg", sheet, [cv2.IMWRITE_JPEG_QUALITY, 94])
    assert ok
    pt_w, pt_h = w / DPI * 72.0, h / DPI * 72.0
    doc = pymupdf.open()
    page = doc.new_page(width=pt_w, height=pt_h)
    page.insert_image(pymupdf.Rect(0, 0, pt_w, pt_h), stream=jpg.tobytes())
    doc.set_metadata({"title": "B240-01200  P&ID No.120-6001  UNIT 1600 EG BLEED RECOVERY",
                      "subject": "分割スキャン (3枚) を1枚に合成", "producer": "", "creator": ""})
    pdf = f"{OUT_DIR}/B240-01200_PID_merged.pdf"
    doc.save(pdf, deflate=True, garbage=4)
    doc.close()

    json.dump({"canvas": [W, H], "frame": [left, int(Y_TOP), right, int(Y_TOP + FRAME_H)],
               "crop": [x0, y0, x1, y1],
               "seams_x_in_crop": [1537 - x0, 4287 - x0],
               "dpi": DPI, "mm": [round(w / DPI * 25.4, 1), round(h / DPI * 25.4, 1)]},
              open(f"{OUT_DIR}/_merge_meta.json", "w"), indent=1)
    print(f"出力: {w} x {h} px = {w / DPI * 25.4:.0f} x {h / DPI * 25.4:.0f} mm @ {DPI:.0f}dpi")
    print(f"  {png}")
    print(f"  {pdf}")


if __name__ == "__main__":
    main()
