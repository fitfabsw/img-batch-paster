#!/usr/bin/env bash
# 在 lab mac mini 上停用並移除 launchd agent
set -euo pipefail
LABEL="com.zealzel.imgbatchpaster"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
launchctl unload "$PLIST" 2>/dev/null || true
rm -f "$PLIST"
echo "✓ 已停用 $LABEL（repo 與 .venv 保留）"
