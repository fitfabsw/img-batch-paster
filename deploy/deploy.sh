#!/usr/bin/env bash
# 從你本機 push 到 GitHub 後跑此腳本：
#   SSH 到 lab mac mini → 切到指定 branch → git pull → pip sync → 重啟服務
#
# 用法：
#   ./deploy/deploy.sh                              # 預設 main
#   BRANCH=feat-A1-fix1 ./deploy/deploy.sh          # 指定 branch
#   SSH_TARGET=cpdx_sw@10.35.36.168 REMOTE_DIR='~/img-batch-paster' ./deploy/deploy.sh
set -euo pipefail

SSH_TARGET="${SSH_TARGET:-cpdx_sw@10.35.36.168}"
REMOTE_DIR="${REMOTE_DIR:-\$HOME/img-batch-paster}"
# 若使用者傳了 "~/..."，轉成 "$HOME/..." 讓遠端 bash 展開
case "$REMOTE_DIR" in "~"*) REMOTE_DIR="\$HOME${REMOTE_DIR#\~}" ;; esac
LABEL="com.zealzel.imgbatchpaster"
PORT="${PORT:-5050}"
BRANCH="${BRANCH:-main}"

echo "→ Deploy to $SSH_TARGET:$REMOTE_DIR  (branch: $BRANCH)"

ssh "$SSH_TARGET" bash -se <<EOF
set -euo pipefail
cd "$REMOTE_DIR"

echo "[1/3] git fetch + checkout $BRANCH + pull…"
git fetch --prune origin
# 若有未 commit 變更先 stash 起來，避免 checkout 失敗
if ! git diff --quiet || ! git diff --cached --quiet; then
    echo "  ⚠ 遠端有未 commit 變更，先 stash"
    git stash push -u -m "auto-stash before deploy \$(date +%Y%m%d-%H%M%S)"
fi
git checkout "$BRANCH"
git pull --rebase origin "$BRANCH"

echo "[2/3] 同步依賴…"
./.venv/bin/pip install --quiet -e .

echo "[3/3] 重啟服務…"
launchctl kickstart -k "gui/\$(id -u)/$LABEL"
sleep 2

if curl -sf -o /dev/null "http://127.0.0.1:$PORT/"; then
    echo "✓ active (branch: \$(git rev-parse --abbrev-ref HEAD) @ \$(git rev-parse --short HEAD))"
else
    echo "✗ 服務未啟動，最後 20 行 err log："
    tail -n 20 /tmp/img-batch-paster.err 2>/dev/null || true
    exit 1
fi
EOF
