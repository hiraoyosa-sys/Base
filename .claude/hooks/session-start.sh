#!/bin/bash
# SessionStart hook（リモート＝Claude Code on the web 用）
# 目的：毎セッション開始時に「GDriveの起動プロトコルを必ず読め」という指示を
#       コンテキストへ注入する（＝リモートだけ抜けていた"最初に読む仕込み"）。
# 注意：フックは shell であり MCP を呼べないため、ここでは取得せず"指示"を出すだけ。
#       実際の取得は CLAUDE.md に従ってエージェントが Google Drive MCP で行う。
set -uo pipefail

# scripts/ の作図・Excel生成に必要な依存を静かに導入（出力品質とは無関係。ログはstderrへ）。
{
  pip install --quiet --disable-pip-version-check pillow xlsxwriter openpyxl >/dev/null 2>&1 || true
} 2>/dev/null

# ここからの stdout はセッションのコンテキストに注入される
cat <<'MSG'
==================================================================
 起動プロトコル（リモート）：作業前に必ず実行すること
==================================================================
このリポジトリはルール・知識の正本ではない。正本は Google Drive の
Obsidian `_Claude/起動プロトコル.md`。作業に入る前に、Google Drive MCP の
download_file_content で次を読み、その指示に従うこと（read_file_content は
.md 非対応で失敗するため使わない）：

  起動プロトコル.md   fileId 1SlSHA4yzvyWcLA5yVizVOn3P6g_Lvxst
  参照プロトコル.md   fileId 1ngeZCK1nJ7LWWJ2i4F0zF9pWOUh2qP9i
  ルール.md           fileId 12zGPfY6Cvb2O3Qg7zZyMoAKZSfioFXEU
  修正ログ.md         fileId 1lC6_8Xplko0y5KkR5bgYHf3QZqV8CA-q
  セッション履歴.md   fileId 1kEYEBlR9xTwanuIhn1_qHDzjuyrNEGyF（末尾を見る）

pptx/Excel/資料を出す直前に「修正ログ.md」を照合してから出力すること。
詳細は CLAUDE.md を参照。
==================================================================
MSG
