#!/usr/bin/env bash
# 跑全部自動化測試：重建 grid fixtures → 用真實前端函式驗證 6 個 xlsx 依檔名 placement case。
# 用法：bash tests/run_tests.sh
set -e
cd "$(dirname "$0")/.."
PY=".venv/bin/python"; [ -x "$PY" ] || PY="python3"
echo "→ 產生 grid fixtures"
"$PY" tests/gen_grids.py
echo "→ placement 測試（真實前端函式）"
node tests/test_placement.mjs
