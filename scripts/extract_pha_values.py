#!/usr/bin/env python3
"""PHA_Organizer 3.1.1 の project.pha から設定値（故障率・デマンド確率）を抽出する。

.pha は先頭8バイトの独自マジックナンバーに続き、phaorg.Project を独自バイナリ形式で
シリアライズしたもの。フィールドは「タグ(uint32) + 型(uint16) + 値」の並びで格納される。
  型 0x07 = 倍精度浮動小数（8バイト）
  型 0x08 = 文字列（int32長 + UTF-8）
  型 0x0b = int32
  型 0x0d = uint16（bool相当）
テンプレート要素は tag 0x01=要素名, 0x24=故障モード, 0x25=時間確率λ(/hr),
0x26=デマンド確率(/d), 0x28=点検周期T(h) を持つ。
検証: 温度計 λ=9.7e-5, T=720 のとき 0x26=0.035 ≒ λ*T/2（一致確認済み 2026-07-21）。

usage: python3 extract_pha_values.py project.pha
"""
import json
import struct
import sys


def tokenize(data: bytes):
    toks = []
    i, n = 0, len(data)
    while i < n - 6:
        tag = struct.unpack_from('<I', data, i)[0]
        typ = struct.unpack_from('<H', data, i + 4)[0]
        if tag < 0x300:
            if typ == 0x08 and i + 10 <= n:
                slen = struct.unpack_from('<I', data, i + 6)[0]
                if 0 <= slen < 2000 and i + 10 + slen <= n:
                    try:
                        s = data[i + 10:i + 10 + slen].decode('utf-8')
                        toks.append((i, tag, 'str', s))
                        i += 10 + slen
                        continue
                    except UnicodeDecodeError:
                        pass
            elif typ == 0x07 and i + 14 <= n:
                toks.append((i, tag, 'dbl', struct.unpack_from('<d', data, i + 6)[0]))
                i += 14
                continue
            elif typ == 0x0b and i + 10 <= n:
                toks.append((i, tag, 'int', struct.unpack_from('<i', data, i + 6)[0]))
                i += 10
                continue
            elif typ == 0x0d and i + 8 <= n:
                toks.append((i, tag, 'bool', struct.unpack_from('<H', data, i + 6)[0]))
                i += 8
                continue
        i += 1
    return toks


def extract_records(toks):
    rows = []
    for idx, (off, tag, kind, val) in enumerate(toks):
        if kind == 'str' and tag == 0x24:  # 故障モード
            names = []
            j = idx - 1
            while j >= 0 and len(names) < 2:
                o, t, k, v = toks[j]
                if k == 'str' and t == 0x01 and v:
                    names.append(v)
                j -= 1
            comp = names[0] if names else ''
            fields = {}
            j = idx + 1
            while j < len(toks):
                o, t, k, v = toks[j]
                if k == 'str' and t in (0x01, 0x24):
                    break
                if t in (0x25, 0x26, 0x27, 0x28, 0x29, 0x2c):
                    fields[t] = v
                j += 1
            rows.append(dict(component=comp, mode=val,
                             lambda_hr=fields.get(0x25), pfd=fields.get(0x26),
                             ptype=fields.get(0x27), T_hr=fields.get(0x28),
                             extra=fields.get(0x29), category=fields.get(0x2c)))
    uniq = {}
    for r in rows:
        key = (r['component'], r['mode'], r['lambda_hr'], r['pfd'])
        uniq.setdefault(key, r)
    return list(uniq.values())


def main():
    path = sys.argv[1]
    data = open(path, 'rb').read()
    recs = extract_records(tokenize(data))
    print(json.dumps(recs, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
