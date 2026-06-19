# CLAUDE.md — リモート（Claude Code on the web）起動ブートストラップ

> このリポジトリは「成果物置き場」であり、ルール・知識の**正本ではない**。
> 正本は Google Drive の Obsidian `_Claude/起動プロトコル.md`（＝全Claude環境が最初に読む司令塔）。
> ローカル/デスクトップは `~/.claude/CLAUDE.md` の `@import` で、Claude.ai は Project 指示で
> 起動プロトコルを読んでいる。**このファイルは、その"リモート版の仕込み"**である。
> （DRY：ルール本文をここに複製しない。必ず GDrive の正本を読みに行く。）

---

## 作業を始める前に必ずやること（例外なし）

1. **Google Drive MCP で `起動プロトコル.md` を読む。**
   - fileId: `1SlSHA4yzvyWcLA5yVizVOn3P6g_Lvxst`
   - 取得は **`download_file_content`** を使う。
     ⚠️ `read_file_content` は `.md`(text/markdown) 非対応で「No approval received」で失敗する（2026-06-18 検証確定）。
2. 起動プロトコルの手順（§1「読の3ループ」）に従い、続けて次を読む（同じ `download_file_content`）：
   - `参照プロトコル.md`（話題→ノートの地図）… fileId `1ngeZCK1nJ7LWWJ2i4F0zF9pWOUh2qP9i`
   - `ルール.md`（核原則）… fileId `12zGPfY6Cvb2O3Qg7zZyMoAKZSfioFXEU`
   - `修正ログ.md`（用語・機器番号・数値の取り違え集＝同じ指摘を二度させない）… fileId `1lC6_8Xplko0y5KkR5bgYHf3QZqV8CA-q`
   - `セッション履歴.md` の**末尾エントリ**（前回どこまで）… fileId `1kEYEBlR9xTwanuIhn1_qHDzjuyrNEGyF`
   - `書込プロトコル.md`（残し方）… fileId `1CyhmdXAK0axd3BNPwwjKzeP4NI2QgRUA`
3. 機器番号(R-1101等)・PJ名(M50/温度計削減等)・専門用語・**資料(pptx/Excel)作成**が出たら、
   まず `参照プロトコル.md` で該当ノートに当たりを付け、その 40_Knowledge 索引を読んでから動く。**推測で断定しない。**

> `_Claude` フォルダ（上記ファイルの親）fileId: `1G4VeVRPHBM8s6uDEODxFnmfBqb8gSpBp`
> 個別 fileId が変わっていたら、このフォルダを `parentId = '...'` で一覧し直して取得する。

---

## 書き戻し（教訓・ミス・進捗を残す）

- リモートからは vault の `.md` を直接編集できない。**`_Claude/inbox/`（fileId `1mVNGFdaqvPkXTh9busOSlo271ky43OU8`）に下書きを `create_file` で投下**する（書込プロトコル準拠）。後でローカルが正本へマージする。
- 新しいミス指摘を受けたら必ず inbox に「修正ログ追補」の下書きを残す。会話だけで終わらせない（次セッションで消える）。

---

## このリポジトリでの作業ルール（最低限）

- 成果物は `outputs/`、生成/作図スクリプトは `scripts/`。Python 生成物パスは `/home/user/Base/outputs`。
- **資料の見た目・表現・用語・グルーピングのルールは、GDrive の `ルール.md`／`修正ログ.md` に従う。**
  （色付き枠・カラフル配色・「→」記号書き・独自記号や造語・抽象語・削りすぎ 等は過去に何度も手戻りになっている。
  出力直前に `修正ログ.md` を照合してから出す。）
- 環境：Web セッションは毎回まっさらなコンテナで**このリポジトリだけ**を clone する。
  だから知識はこのリポジトリに無く、GDrive の正本を読みに行く必要がある（上記手順）。
