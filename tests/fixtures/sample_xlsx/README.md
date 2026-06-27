# xlsx 依檔名 測試範例

最簡測試集：4 張圖 + 6 個範本，涵蓋 横式/直式 × (全空 / 只有 index / index+group)。

## 測試圖（`images/`）

檔名 = `{group}-{index}`：

```
AAA-1.png   AAA-2.png   BBB-2.png   BBB-3.png
```

→ Group = {AAA, BBB}、Index = {1, 2, 3}

## 範本（`templates/`）

| 檔案 | 方向 | 範本內容 | 自動偵測 |
|---|---|---|---|
| `h1_empty.xlsx` | 横式 | 全空（只有框線） | 横式 |
| `h2_index.xlsx` | 横式 | 頂列 index：2,3,4 | 横式 |
| `h3_index_group.xlsx` | 横式 | 頂列 index：2,3,4 ＋ 左欄 group：BBB,CCC | 横式 |
| `v4_empty.xlsx` | 直式 | 全空（只有框線） | ⚠ 横式（需手選「直式」） |
| `v5_index.xlsx` | 直式 | 左欄 index：2,3,4 | 直式 |
| `v6_index_group.xlsx` | 直式 | 頂列 group：BBB,CCC ＋ 左欄 index：2,3,4 | 直式 |
| `h7_group.xlsx` | 横式 | 左欄 group：BBB,CCC（無 index） | 横式 |
| `v8_group.xlsx` | 直式 | 頂列 group：BBB,CCC（無 index） | 直式 |

横式＝Group 在左欄、Index 在頂列；直式＝相反。

## 預期貼圖結果（依檔名）

- **全空（h1 / v4）**：Group、Index 都「依檔名」自動產生並寫入。4 張全貼（2 group × 各自 index）。
- **只有 index（h2 / v5）**：Index「依範本」(2,3,4)；Group「依檔名」(AAA,BBB) 自動產生。
  - `AAA-1` 的 index=1 不在範本 → **跳過**。其餘 3 張貼入。
- **index+group（h3 / v6）**：兩軸都「依範本」。
  - Group 只有 BBB,CCC → **AAA 整組跳過**；Index 只有 2,3,4 → index=1 跳過。
  - 結果：只有 `BBB-2`、`BBB-3` 貼入（對到 BBB 列/欄 × index 2,3）。
- **只有 group（h7 / v8）**：Group「依範本」、Index「依檔名」。
  - Group 只有 BBB,CCC → **AAA 整組跳過**；Index 依檔名自動產生。
  - 結果：`BBB-2`、`BBB-3` 貼入（方向靠 group 軸命中自動判定，h7=横式、v8=直式）。

## 注意

- **v4（直式全空）無法自動偵測**：沒有任何標籤可比對方向 → 會落回横式，請在「範本方向」手選**直式**。
- 重新產生：`.venv/bin/python tests/fixtures/sample_xlsx/make_samples.py`
