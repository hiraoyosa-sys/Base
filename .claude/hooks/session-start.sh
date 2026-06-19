#!/bin/bash
# SessionStart hook: 第二の脳・起動ブートストラップをセッション開始時に機械注入する。
# 平尾の手動指示なしで、リモート Claude Code が起動プロトコルを自動読込するための仕組み。
set -euo pipefail

# このスクリプトの場所から正本ファイルを解決（cwd 非依存）
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BOOT_FILE="$SCRIPT_DIR/../second-brain-boot.md"

# 正本が無ければ最低限の指示を注入（フォールバック）
if [[ -f "$BOOT_FILE" ]]; then
  CONTEXT="$(cat "$BOOT_FILE")"
else
  CONTEXT="【第二の脳・起動】本作業前に Google Drive MCP の download_file_content で 起動プロトコル.md (fileId 1SlSHA4yzvyWcLA5yVizVOn3P6g_Lvxst) を読み、その §1〜§5 に従うこと。read_file_content は .md 非対応のため使わない。"
fi

# SessionStart の additionalContext として注入（JSON は jq で安全に生成）
jq -n --arg ctx "$CONTEXT" \
  '{hookSpecificOutput: {hookEventName: "SessionStart", additionalContext: $ctx}}'
