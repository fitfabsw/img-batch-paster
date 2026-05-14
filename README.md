# img-batch-paster

把一個資料夾裡的圖片，批次貼到 PowerPoint / Keynote 投影片的網格上。

服務在 lab Mac mini：**http://10.35.36.168:5050/**

## Web 使用方式（建議）

1. **範本** — 點「📤 上傳範本…」選自己的 `.pptx`，或按「預設」用內建範本
2. **圖片** — 點「📤 上傳圖片資料夾…」整個資料夾上傳
3. **網格** — 設定第一張圖位置、圖寬、間距、每頁 cols × rows（左側即時調整、右側預覽）
4. **檔案順序** — 中欄拖 ▲▼ 重排，或點清單項目跳到對應頁
5. **匯出** — 選 `.pptx` 或 `.key`，按按鈕，瀏覽器自動下載

預設：欄 3 × 列 3，每頁 9 張；超過自動分頁。圖片高度依原始比例自動算，不會被拉伸。

### 限制

- `.key` 匯出只能在 mac mini 跑（依賴 Keynote.app）
- 中欄上傳資料夾用瀏覽器原生「webkitdirectory」，Chrome / Safari 支援良好
- 暫存檔放 server `/tmp/img-batch-paster-uploads/<session>/`，會隨重開機清掉

### 本機開發

```bash
python3 -m venv .venv && .venv/bin/pip install -e .
.venv/bin/img-batch-paster-web        # → http://127.0.0.1:5050/
```

部署到 lab mac mini 請看 [deploy/SETUP.md](deploy/SETUP.md)。

---

## CLI 使用方式（舊版，已被 Web 取代）

最早的版本是 CLI + YAML 設定檔，目前仍可用。

### 安裝

```bash
python3 -m venv .venv
.venv/bin/pip install -e .
```

### 使用

1. 複製設定範本：
   ```bash
   cp config.example.yaml my-config.yaml
   ```
2. 編輯 `my-config.yaml`，設定資料夾、欄數、網格座標／大小／間距、輸出路徑。
3. 執行：
   ```bash
   .venv/bin/img-batch-paster -c my-config.yaml
   ```
   可選參數：
   - `--folder PATH` 覆蓋 `input.folder`
   - `-o PATH` 覆蓋 `output.path`

### 設定欄位

```yaml
slide:                  # 投影片尺寸 (cm)
  width_cm: 25.4
  height_cm: 14.29
grid:
  origin: {x_cm, y_cm}  # 第一格左上角
  cell:   {w_cm, h_cm}  # 每格大小
  gap:    {x_cm, y_cm}  # 格與格間距
  cols:   3             # 欄數 (列數依檔名自動)
input:
  folder: ./images
  pattern: "{group}_{n}"   # 同 group 為一列，n 為欄編號
  extensions: [".png", ".jpg", ".jpeg"]
output:
  path: ./out.pptx
  template: null        # 可選的 .pptx 範本
```

範例：`xxx_1.png, xxx_2.png, xxx_3.png, yyy_1.png, yyy_2.png, yyy_3.png` → 2 列 × 3 欄。

### 測試資料

```bash
.venv/bin/python tests/make_test_images.py
.venv/bin/img-batch-paster -c tests/fixtures/config.yaml \
  --folder tests/fixtures/images -o tests/fixtures/out.pptx
```

---

## 路線圖

- [x] M1：CLI + YAML 設定 + `.pptx` 輸出
- [x] M2：Keynote 輸出（AppleScript）
- [x] M3：Web UI 即時預覽（Flask + React CDN）
- [x] M4：超出列數自動多頁
- [x] M5：瀏覽器上傳檔案 + 自動下載
- [x] M6：部署到 lab Mac mini（launchd）

## 技術棧

Python 3.10+ / python-pptx / Pillow / PyYAML / click / Flask / React (CDN) / Tailwind (CDN) / LibreOffice (預覽渲染) / Keynote.app + AppleScript (.key 匯出)
