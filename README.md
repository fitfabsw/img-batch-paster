# img-batch-paster

把資料夾內的圖片依檔名規則，批次貼到簡報（PowerPoint / Keynote）的表格網格中。

## 功能（M1 — CLI）

- 掃描指定資料夾的圖片
- 依檔名 pattern `{group}_{n}` 分組：
  - 同 `group` 為**同一列**
  - `n` 為**欄編號**（1-based）
- 依 YAML 設定的網格位置、大小、間距，貼到一張投影片中
- 輸出 `.pptx`

範例：`xxx_1.png, xxx_2.png, xxx_3.png, yyy_1.png, yyy_2.png, yyy_3.png` → 2 列 × 3 欄。

## 安裝

```bash
python3 -m venv .venv
.venv/bin/pip install -e .
```

## 使用

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

## 設定欄位

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
  pattern: "{group}_{n}"
  extensions: [".png", ".jpg", ".jpeg"]
output:
  path: ./out.pptx
  template: null        # 可選的 .pptx 範本
```

## 測試

```bash
.venv/bin/python tests/make_test_images.py
.venv/bin/img-batch-paster -c tests/fixtures/config.yaml \
  --folder tests/fixtures/images -o tests/fixtures/out.pptx
```

## 路線圖

- [x] M1：CLI + YAML 設定 + `.pptx` 輸出
- [ ] M2：Keynote 輸出（AppleScript）
- [ ] M3：Web UI 即時預覽（PyWebView + Flask + React）
- [ ] M4：超出列數自動多張投影片

## 技術栈

Python 3.10+ / python-pptx / Pillow / PyYAML / click
