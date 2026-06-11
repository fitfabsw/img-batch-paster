# img-batch-paster

把資料夾裡的圖片批次貼到 PowerPoint / Keynote / Excel 的固定格子。Web UI 為主，CLI 為輔。

## Stack

- **Backend**: Python 3.10 + Flask（內建 dev server）
- **Frontend**: React (no-build) — 整個 app 寫在 `static/index.html`，瀏覽器即時 Babel 編譯
- **Type**: single-process（Flask 同時 serve API + static）
- **Port**: 5050（見 ~/.claude/skills/infra/dev/ports.md）
- **Deps manager**: pip（`pyproject.toml`，editable install）— uv 亦可
- **Test**: 無單元測試；`tests/make_test_images.py` 產測試圖、`tests/templates/` 放範本
- **Deploy target**: lab-mac-mini via launchd（再經 it-server 反代對外）
- **Entry**: `src/img_batch_paster/web/app.py`（`img-batch-paster-web` entry point）

## 重點目錄

- `src/img_batch_paster/web/app.py` — Flask routes（`/api/*`：上傳、預覽、配置、匯出）
- `src/img_batch_paster/web/static/index.html` — 全部前端 UI（單檔 ~3267 行，含 A1/A2 模式邏輯）
- `src/img_batch_paster/web/template_render.py` — 範本首頁渲染（Keynote 優先、LibreOffice fallback）
- `src/img_batch_paster/keynote_export.py` — Keynote 匯出 / 渲染（AppleScript 包裝，僅 macOS）
- `src/img_batch_paster/{pptx_writer,xlsx_writer}.py` — .pptx / .xlsx 輸出
- `deploy/` — launchd plist + install/deploy/uninstall scripts

## 跑起來

```bash
python3 -m venv .venv && .venv/bin/pip install -e .
.venv/bin/img-batch-paster-web    # → http://127.0.0.1:5050/
```

## 測試

```bash
.venv/bin/python tests/make_test_images.py    # 產生測試圖到 tests/fixtures/
# 無自動化測試，靠手動跑 web UI 驗證
```

## Deploy

lab-mac-mini via launchd plist，見 [deploy/SETUP.md](deploy/SETUP.md)。
另經 it-server `/imgpaste/` 反代對外，見 [infra recipe](~/.claude/skills/infra/recipes/expose-lab-mac-mini-service-via-it-server.md)。

## 注意

- 前端是 no-build：改 JSX 直接編 `static/index.html`，**不需** `npm install` / build step。瀏覽器跑 development build of React（Babel runtime 編譯），效能可接受但非最佳。
- Port 5050 **寫死於** `deploy/com.zealzel.imgbatchpaster.plist` + `deploy/install.sh`，不可隨意改（it-server nginx 也指過去）。
- **macOS TCC 路徑限制**：launchd 跑的 process **不能存取** `~/Documents/`、`~/Desktop/`、`~/Downloads/`，repo 必須放 `~/` 第一層。
- **Keynote + launchd 的 TCC 坑**：launchd-spawned 的 Python 沒有「人類祖先」，第一次控制 Keynote（.key 匯出 / 範本渲染）的 Automation 授權對話框跳不出來 → AppleScript 卡住 → nginx 504 / 前端 `Unexpected token '<'`。解法：先從互動式 terminal（ssh + tmux）跑一次 dev server，手動點「允許」後，launchd 才吃得到授權。
- **Keynote 14.2 AppleScript bug**：`count of documents` 會回 -1708 失敗。若改 keynote_export.py 的等待邏輯要避開 collection-access API。
- **uv 自帶 Python 的 TCC 問題**：別用 uv 下載的 python-build-standalone（macOS TCC 不認其簽章，會讓 osascript 控 Keynote 失敗）；用系統 / Homebrew / pyenv 的 Python。
- 分支：`feat-A` 為交付主線（→ merge main）；`feat-mcp` = feat-A + MCP server（`mcp_server.py` + `paste_job.py`），web UI 兩邊一致。
```
