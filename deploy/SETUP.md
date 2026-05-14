# Deploy 到 lab Mac mini

目標：`cpdx_sw@10.35.36.168`，launchd user agent，LAN 同事可連。

## 一次性安裝（在 lab mac mini 上）

```bash
# 1. 把 SSH 公鑰加到 mac mini（方便之後免密碼 deploy）
ssh-copy-id cpdx_sw@10.35.36.168

# 2. SSH 進 mac mini
ssh cpdx_sw@10.35.36.168

# 3. 安裝必要工具
brew install --cask libreoffice         # 範本預覽用 soffice
# Keynote.app 從 App Store 安裝（.key 匯出用）

# 4. clone 並執行 install
git clone <repo-url> ~/img-batch-paster
cd ~/img-batch-paster
./deploy/install.sh
```

`install.sh` 會：
- 建立 `.venv`、`pip install -e .`
- 將 launchd plist 生成到 `~/Library/LaunchAgents/com.zealzel.imgbatchpaster.plist`
- `launchctl load` 啟動服務
- 健康檢查並印出 LAN 網址

## 授權自動化權限（第一次 .key 匯出時）

第一次匯出 `.key` 時 macOS 會彈視窗詢問是否允許 `python3` / `bash` 控制 Keynote。
若沒有跳出對話框，可手動加：

```
系統設定 → 隱私權與安全性 → 自動化 → 啟用 Python（或 Terminal）→ Keynote
```

## 更新（在你本機）

push 到 GitHub 後：

```bash
./deploy/deploy.sh
```

預設目標 `cpdx_sw@10.35.36.168:~/img-batch-paster`；覆蓋：

```bash
SSH_TARGET=cpdx_sw@10.35.36.168 REMOTE_DIR='~/img-batch-paster' ./deploy/deploy.sh
```

腳本會在遠端 `git pull` → `pip install -e .` → `launchctl kickstart -k` 重啟 → 健康檢查。

## 操作指令（在 mac mini 上）

| 動作 | 指令 |
|---|---|
| 重啟 | `launchctl kickstart -k gui/$(id -u)/com.zealzel.imgbatchpaster` |
| 停止 | `launchctl unload ~/Library/LaunchAgents/com.zealzel.imgbatchpaster.plist` |
| 啟用 | `launchctl load   ~/Library/LaunchAgents/com.zealzel.imgbatchpaster.plist` |
| log    | `tail -f /tmp/img-batch-paster.log` |
| err    | `tail -f /tmp/img-batch-paster.err` |
| 完全移除 | `./deploy/uninstall.sh` |

## 設定

- 預設 port **5050**，bind `0.0.0.0`（LAN 同事可連 `http://10.35.36.168:5050/`）
- 改 port：編輯 `deploy/com.zealzel.imgbatchpaster.plist` 的 `--port`，重跑 `./deploy/install.sh`
- 開機自動啟動：launchd 內建（`RunAtLoad=true`）
- 崩潰自動重啟：launchd 內建（`KeepAlive=true`）

## 注意

- macOS 防火牆若開啟，第一次外網連線時會詢問是否允許 `python3` 接受連線
- mac mini 進入睡眠時服務不可用；建議到「節能」設定關閉自動睡眠

## 架構備忘：為何不用 Caddy？

cpdx-ai-hub 有用 Caddy，但這個專案不需要：

| | cpdx-ai-hub | img-batch-paster |
|---|---|---|
| 後端 | FastAPI + uvicorn `:8003`（localhost-only） | Flask `:5050` 直接 bind `0.0.0.0` |
| 前端 | React + Vite **打包**成 `dist/`，需要 web server | React 走 **CDN**，由 Flask static 路由直接吐 `index.html` |
| 對外 | Caddy 80/443 同時服務 dist/ + reverse proxy `/api/*` 到 :8003 | 不需要，Flask 單一 process 全包 |

我們的部署是「Flask 一個 process 同時提供 API + 前端」，使用者直接連 `http://<mac-mini-ip>:5050/`。

**未來需要加 Caddy 的情境**：
- 想用 HTTPS（自簽或 Tailscale TLS）
- 想對外開 80/443 port 不寫埠號
- 想加帳號密碼 / IP 白名單
- 想把多個內部小工具集中到同一個網域底下（如 `tools.lab/img-paster/`）
