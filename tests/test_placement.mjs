// Excel 依檔名 placement 測試：用 sample_xlsx fixtures 跑「真實的」前端 placement 函式。
// 直接從 index.html 抽出純函式（colLetterToIdx..buildRows）在 node 執行，不重寫邏輯。
//   先跑 gen_grids.py 產 grids.json。執行：node tests/test_placement.mjs
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import assert from "assert";
import { get_column_letter } from "./_colletter.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

// ── 抽出 index.html 的純 placement 函式並 eval ──
const html = fs.readFileSync(path.join(ROOT, "src/img_batch_paster/web/static/index.html"), "utf8");
const script = html.match(/<script type="text\/babel"[^>]*>([\s\S]*?)<\/script>/)[1];
const start = script.indexOf("function colLetterToIdx");
const end = script.indexOf("function computePages");
assert(start > 0 && end > start, "找不到 placement 函式區塊");
const block = script.slice(start, end);
const exported = ["detectExcelTable", "readAxisLabels", "resolveExcelOrientation",
  "computeExcelCellsAuto", "computeExcelCellsTransposed", "computeExcelCellsHorizGroupTemplate"];
const F = new Function(block + "\nreturn {" + exported.join(",") + "};")();

const grids = JSON.parse(fs.readFileSync(path.join(ROOT, "tests/fixtures/sample_xlsx/grids.json"), "utf8"));
const FILES = ["AAA-1", "AAA-2", "BBB-2", "BBB-3"].map((n) => ({ name: n + ".png", path: n + ".png" }));

// 模擬 UI 對每個範本自動解析的設定（範本方向 + Group/Index 對位）
function configFor(grid, forceOrient) {
  const d = F.detectExcelTable(grid);
  const excel = { startCell: d.startCell, snCol: d.snCol, cellCols: 1, cellRows: 1, gapRows: 0,
    orient: forceOrient || "auto" };
  const label = { pattern: "{group}-{idx}", idxSort: "auto", idxOrder: [], groupSrc: "auto",
    idxIgnore: [], font_pt: 12 };
  const orientation = F.resolveExcelOrientation(grid, excel, label, FILES, true);
  const ax = F.readAxisLabels(grid, d);
  const idxAxis = orientation === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr;
  if (idxAxis.length) { label.idxSort = "custom"; label.idxOrder = [...idxAxis]; }  // Index 依範本
  return { excel, label };
}

// cells → { 檔名: "C3" }（只取圖片 cell）
function placements(grid, forceOrient) {
  const { excel, label } = configFor(grid, forceOrient);
  const cells = F.computeExcelCellsAuto(grid, excel, label, FILES, true);
  const out = {};
  for (const c of cells) {
    if (!c.path) continue;
    const stem = c.path.replace(/\.[^.]+$/, "");
    out[stem] = get_column_letter(c.col) + c.row;
  }
  return out;
}

// 預期結果（依 README；橫式 vs 直式 = 轉置）
const EXPECT = {
  // 全空：兩軸依檔名，4 張全貼（横式資料從表格頂列起、直式從頂列下一列起）
  h1_empty:        { "AAA-1": "C2", "AAA-2": "D2", "BBB-2": "D3", "BBB-3": "E3" },
  v4_empty:        { "AAA-1": "C3", "AAA-2": "C4", "BBB-2": "D4", "BBB-3": "D5" },
  // 只有 index(2,3,4)：Index 依範本 → idx=1(AAA-1) 跳過；Group 依檔名(AAA 列, BBB 列)
  h2_index:        { "AAA-2": "C3", "BBB-2": "C4", "BBB-3": "D4" },
  v5_index:        { "AAA-2": "C3", "BBB-2": "D3", "BBB-3": "D4" },
  // index+group：兩軸依範本 → 只有 BBB 對到（AAA 整組、idx=1 皆跳過）
  h3_index_group:  { "BBB-2": "C3", "BBB-3": "D3" },
  v6_index_group:  { "BBB-2": "C3", "BBB-3": "C4" },
  // 只有 group(BBB,CCC)：Group 依範本 → AAA 跳過；Index 依檔名（方向靠 group 軸命中自動判定）
  h7_group:        { "BBB-2": "D3", "BBB-3": "E3" },
  v8_group:        { "BBB-2": "C4", "BBB-3": "C5" },
};
// v4 全空無法自動偵測直式 → 手選 vertical
const FORCE = { v4_empty: "vertical" };

// 手動覆寫設定的回歸案例：{ grid, orient, groupSrc, idxSrc('template'|'filename'), expect }
function placementsManual(gridName, { orient, groupSrc, idxSrc }) {
  const grid = grids[gridName];
  const d = F.detectExcelTable(grid);
  const excel = { startCell: d.startCell, snCol: d.snCol, cellCols: 1, cellRows: 1, gapRows: 0, orient };
  const ax = F.readAxisLabels(grid, d);
  const idxAxis = orient === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr;
  const label = { pattern: "{group}-{idx}", idxSort: "auto", idxOrder: [], groupSrc, idxIgnore: [], font_pt: 12 };
  if (idxSrc === "template") { label.idxSort = "custom"; label.idxOrder = [...idxAxis]; }
  const cells = F.computeExcelCellsAuto(grid, excel, label, FILES, true);
  const out = {};
  for (const c of cells) if (c.path) out[c.path.replace(/\.[^.]+$/, "")] = get_column_letter(c.col) + c.row;
  return out;
}
const MANUAL = {
  // B-022 回歸：h7（group 在左欄）+ Group/Index 都依檔名 → 資料須從表頭下一列(第3列)起，不騎到範本標籤
  "h7 group=檔名,index=檔名": { gridName: "h7_group", cfg: { orient: "horizontal", groupSrc: "filename", idxSrc: "filename" },
    expect: { "AAA-1": "C3", "AAA-2": "D3", "BBB-2": "D4", "BBB-3": "E4" } },
};

let fail = 0;
for (const name of Object.keys(EXPECT)) {
  const got = placements(grids[name], FORCE[name]);
  try {
    assert.deepStrictEqual(got, EXPECT[name]);
    console.log(`  ✓ ${name}`);
  } catch (e) {
    fail++;
    console.log(`  ✗ ${name}\n      expected ${JSON.stringify(EXPECT[name])}\n      got      ${JSON.stringify(got)}`);
  }
}
for (const [name, t] of Object.entries(MANUAL)) {
  const got = placementsManual(t.gridName, t.cfg);
  try {
    assert.deepStrictEqual(got, t.expect);
    console.log(`  ✓ ${name}`);
  } catch (e) {
    fail++;
    console.log(`  ✗ ${name}\n      expected ${JSON.stringify(t.expect)}\n      got      ${JSON.stringify(got)}`);
  }
}
console.log(fail ? `\n${fail} case(s) FAILED` : "\nAll cases passed ✓");
process.exit(fail ? 1 : 0);
