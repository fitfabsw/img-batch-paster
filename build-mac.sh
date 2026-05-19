#!/usr/bin/env bash
# 打包 macOS .app
#   ./build-mac.sh
set -euo pipefail
cd "$(dirname "$0")"

if [ ! -d .venv ]; then
    python3 -m venv .venv
    .venv/bin/pip install --upgrade pip
fi

# 安裝執行期 + 打包工具
.venv/bin/pip install -e . pywebview pyinstaller >/dev/null

# 清理舊輸出
rm -rf build dist

.venv/bin/pyinstaller build_app.spec --clean --noconfirm

echo
echo "✓ Built: dist/img-batch-paster.app"
echo "  Open: open dist/img-batch-paster.app"
echo "  Notes:"
echo "    - 範本預覽背景需要 LibreOffice (brew install --cask libreoffice)"
echo "    - .key 匯出需要 Keynote (App Store 免費)"
