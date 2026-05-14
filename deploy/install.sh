#!/usr/bin/env bash
# 在 lab mac mini 上第一次安裝：
#   git clone <repo> ~/img-batch-paster
#   cd ~/img-batch-paster
#   ./deploy/install.sh
set -euo pipefail

REPO="$(cd "$(dirname "$0")/.." && pwd)"
LABEL="com.zealzel.imgbatchpaster"
PLIST_SRC="$REPO/deploy/$LABEL.plist"
PLIST_DST="$HOME/Library/LaunchAgents/$LABEL.plist"
PORT="${PORT:-5050}"

echo "[1/5] 檢查相依工具…"
command -v python3 >/dev/null || { echo "缺少 python3"; exit 1; }
command -v soffice >/dev/null || echo "  ⚠️  缺少 soffice (LibreOffice)；範本預覽會失敗，請 brew install --cask libreoffice"
[ -d "/Applications/Keynote.app" ] || echo "  ⚠️  未找到 Keynote.app；.key 匯出將無法使用"

echo "[2/5] 建立 venv 並安裝依賴…"
[ -d "$REPO/.venv" ] || python3 -m venv "$REPO/.venv"
"$REPO/.venv/bin/pip" install --quiet --upgrade pip
"$REPO/.venv/bin/pip" install --quiet -e "$REPO"

echo "[3/5] 產生 launchd plist…"
mkdir -p "$HOME/Library/LaunchAgents"
sed "s|__REPO__|$REPO|g" "$PLIST_SRC" > "$PLIST_DST"

echo "[4/5] 載入服務…"
launchctl unload "$PLIST_DST" 2>/dev/null || true
launchctl load "$PLIST_DST"

echo "[5/5] 等待 server 起來…"
for i in $(seq 1 15); do
    if curl -sf -o /dev/null "http://127.0.0.1:$PORT/"; then
        IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "<mac-mini-ip>")
        echo "✓ 服務已啟動"
        echo "  本機: http://127.0.0.1:$PORT/"
        echo "  LAN : http://$IP:$PORT/"
        echo
        echo "Log:    tail -f /tmp/img-batch-paster.log"
        echo "Err:    tail -f /tmp/img-batch-paster.err"
        echo "Stop:   launchctl unload $PLIST_DST"
        exit 0
    fi
    sleep 1
done

echo "✗ 等待 15s 仍未啟動，請看 /tmp/img-batch-paster.err"
exit 1
