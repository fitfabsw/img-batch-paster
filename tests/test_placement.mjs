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
  "computeExcelCellsAuto", "computeExcelCellsTransposed", "computeExcelCellsHorizGroupTemplate",
  "extractGroupIdx", "colLetterToIdx", "idxSrcShown", "detectIdxList"];
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
  // *1 全空：兩軸依檔名，4 張全貼（頂列保留給 index 標頭、資料從下一列起）
  h1_empty:        { "AAA-1": "C3", "AAA-2": "D3", "BBB-2": "D4", "BBB-3": "E4" },
  v1_empty:        { "AAA-1": "C3", "AAA-2": "C4", "BBB-2": "D4", "BBB-3": "D5" },
  // *2 只有 index(2,3,4)：Index 依範本 → idx=1(AAA-1) 跳過；Group 依檔名(AAA 列, BBB 列)
  h2_index:        { "AAA-2": "C3", "BBB-2": "C4", "BBB-3": "D4" },
  v2_index:        { "AAA-2": "C3", "BBB-2": "D3", "BBB-3": "D4" },
  // *3 index+group：兩軸依範本 → 只有 BBB 對到（AAA 整組、idx=1 皆跳過）
  h3_index_group:  { "BBB-2": "C3", "BBB-3": "D3" },
  v3_index_group:  { "BBB-2": "C3", "BBB-3": "C4" },
  // *4 只有 group(BBB,CCC)：Group 依範本 → AAA 跳過；Index 依檔名（方向靠 group 軸命中自動判定）
  h4_group:        { "BBB-2": "D3", "BBB-3": "E3" },
  v4_group:        { "BBB-2": "C4", "BBB-3": "C5" },
};
// v1 全空無法自動偵測直式 → 手選 vertical
const FORCE = { v1_empty: "vertical" };

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
  // B-022 回歸：h4（group 在左欄）+ Group/Index 都依檔名 → 資料須從表頭下一列(第3列)起，不騎到範本標籤
  "h4 group=檔名,index=檔名": { gridName: "h4_group", cfg: { orient: "horizontal", groupSrc: "filename", idxSrc: "filename" },
    expect: { "AAA-1": "C3", "AAA-2": "D3", "BBB-2": "D4", "BBB-3": "E4" } },
  // B-023 回歸：v4(直式,只有 group) + Index=依範本(範本無 index) → 應全空(横式 h4 已是空，直式須一致)
  "v4 index=依範本(無index)": { gridName: "v4_group", cfg: { orient: "vertical", groupSrc: "template", idxSrc: "template" },
    expect: {} },
  "h4 index=依範本(無index)": { gridName: "h4_group", cfg: { orient: "horizontal", groupSrc: "template", idxSrc: "template" },
    expect: {} },
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
// ── 完整矩陣 + 不變量：8 範本 × Group{範本/檔名} × Index{範本/檔名}，自動抓結構性 bug ──
//   不必逐一手算每格，靠「鐵則」覆蓋 off-by-one / 依範本越界 / 空軸不該貼 / 撞格 等整類問題。
function computeFull(gridName, orient, groupSrc, idxSrc) {
  const grid = grids[gridName];
  const d = F.detectExcelTable(grid);
  const ax = F.readAxisLabels(grid, d);
  const excel = { startCell: d.startCell, snCol: d.snCol, cellCols: 1, cellRows: 1, gapRows: 0, orient };
  const idxAxisArr = orient === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr;
  const groupLabels = new Set([...(orient === "vertical" ? ax.topLabels : ax.leftLabels).keys()]);
  const idxLabels = new Set([...(orient === "vertical" ? ax.leftLabels : ax.topLabels).keys()]);
  const label = { pattern: "{group}-{idx}", idxSort: "auto", idxOrder: [], groupSrc, idxIgnore: [], font_pt: 12 };
  if (idxSrc === "template") { label.idxSort = "custom"; label.idxOrder = [...idxAxisArr]; }
  else { label.idxSort = "custom"; label.idxOrder = F.colLetterToIdx ? [] : []; label.idxSort = "auto"; }  // filename = auto
  const cells = F.computeExcelCellsAuto(grid, excel, label, FILES, true);
  const images = cells.filter((c) => c.path).map((c) => {
    const { group, idx } = F.extractGroupIdx(c.path, label.pattern);
    return { row: c.row, col: c.col, group: String(group).toLowerCase(), idx: String(idx).toLowerCase() };
  });
  // 從「範本實際標籤位置」獨立推第一個資料列(不靠 detectExcelTable，才能抓它自己的 off-by-one)：
  //   有左欄標籤 → 從首個左欄標籤列起；否則有頂列標籤 → minR+1(表頭佔一列)；全空 → minR。
  const minR = ax.region.minR, minC = ax.region.minC;
  const textCells = (grid.cells || []).filter((c) => String(c.text ?? "").trim());
  const leftRows = textCells.filter((c) => c.c === minC).map((c) => c.r);
  const hasTopLabel = textCells.some((c) => c.c > minC);
  const dataRowFloor = leftRows.length ? Math.min(...leftRows) : (minR + (hasTopLabel ? 1 : 0));
  return { images, ax, region: ax.region, groupLabels, idxLabels, dataRowFloor };
}
const FORCE_ORIENT = (name) => (name[0] === "v" ? "vertical" : "horizontal");
let matrixFail = 0, matrixRun = 0;
for (const name of Object.keys(grids)) {
  const orient = FORCE_ORIENT(name);
  for (const groupSrc of ["template", "filename"]) {
    for (const idxSrc of ["template", "filename"]) {
      matrixRun++;
      const tag = `${name} [g=${groupSrc[0]},i=${idxSrc[0]}]`;
      let r;
      try { r = computeFull(name, orient, groupSrc, idxSrc); }
      catch (e) { matrixFail++; console.log(`  ✗ ${tag} 例外: ${e.message}`); continue; }
      const { images, region, groupLabels, idxLabels, dataRowFloor } = r;
      const minC = region.minC;
      const errs = [];
      // I1：圖片不得落在資料區之上(表頭列)或標籤欄(off-by-one / 騎到表頭)。dataRowFloor 由範本標籤獨立推得。
      for (const im of images) if (im.row < dataRowFloor || im.col <= minC) errs.push(`圖 ${im.group}-${im.idx} 落在表頭/標籤(${im.col},${im.row})，資料應從第 ${dataRowFloor} 列起`);
      // I2：依範本軸 → 貼進去的值必須是範本既有標籤（不越界、不覆寫無關列欄）
      if (groupSrc === "template") for (const im of images) if (!groupLabels.has(im.group)) errs.push(`group 依範本但貼了範本沒有的 ${im.group}`);
      if (idxSrc === "template") for (const im of images) if (!idxLabels.has(im.idx)) errs.push(`index 依範本但貼了範本沒有的 ${im.idx}`);
      // I3：依範本但該軸範本無標籤 → 不該有任何圖
      if (groupSrc === "template" && groupLabels.size === 0 && images.length) errs.push(`group 依範本+範本無 group，卻貼了 ${images.length} 張`);
      if (idxSrc === "template" && idxLabels.size === 0 && images.length) errs.push(`index 依範本+範本無 index，卻貼了 ${images.length} 張`);
      // I4：不得兩圖同格
      const seen = new Set();
      for (const im of images) { const k = im.col + "," + im.row; if (seen.has(k)) errs.push(`兩圖同格 ${k}`); seen.add(k); }
      if (errs.length) { matrixFail++; console.log(`  ✗ ${tag}\n      ${errs.join("\n      ")}`); }
    }
  }
}
console.log(`  矩陣不變量：${matrixRun - matrixFail}/${matrixRun} 組合通過`);
fail += matrixFail;

// ── UI 狀態：Index 對位「按鈕顯示來源」必須＝placement 實際用的來源（B-021 那類顯示≠行為）──
//   獨立算「實際生效來源」：用真正的 idx 解析結果，對照是否＝範本標籤；再比對 idxSrcShown 的顯示。
function effectiveIdxSrc(gridName, orient, label) {
  const grid = grids[gridName];
  const d = F.detectExcelTable(grid);
  const ax = F.readAxisLabels(grid, d);
  const idxAxis = (orient === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr) || [];
  let resolved;
  if (orient === "vertical") resolved = F.computeExcelCellsTransposed(grid, d, { cellCols: 1, cellRows: 1, gapRows: 0 }, label, FILES).idxList || [];
  else resolved = F.detectIdxList(FILES, label.pattern, label.idxSort || "auto", label.idxOrder || []);
  const same = (a, b) => a.length === b.length && a.every((x, i) => String(x) === String(b[i]));
  // 生效＝依範本：解析結果就是範本 idx 標籤（且明確選了 custom＝範本）；其餘＝依檔名
  const choseTemplate = (label.idxSort === "custom") && same(label.idxOrder || [], idxAxis);
  return choseTemplate ? "template" : "filename";
}
let uiFail = 0, uiRun = 0;
for (const name of Object.keys(grids)) {
  const orient = FORCE_ORIENT(name);
  const d = F.detectExcelTable(grids[name]); const ax = F.readAxisLabels(grids[name], d);
  const idxAxis = (orient === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr) || [];
  const fileOrder = F.detectIdxList(FILES, "{group}-{idx}", "auto", []);
  // 涵蓋三種 label 狀態：auto 預設(B-021 案發點) / 依範本 / 依檔名
  const states = [
    { idxSort: "auto", idxOrder: [] },
    { idxSort: "custom", idxOrder: [...idxAxis] },
    { idxSort: "custom", idxOrder: [...fileOrder] },
  ];
  for (const s of states) {
    uiRun++;
    const label = { pattern: "{group}-{idx}", ...s };
    const shown = F.idxSrcShown(label, idxAxis);
    const eff = effectiveIdxSrc(name, orient, label);
    if (shown !== eff) { uiFail++; console.log(`  ✗ ${name} idxSort=${s.idxSort},order=[${s.idxOrder}] 顯示=${shown} 實際=${eff}`); }
  }
}
console.log(`  UI 顯示一致：${uiRun - uiFail}/${uiRun} 狀態通過`);
fail += uiFail;

// ── 對位鎖定政策：範本有定義該軸→禁「依檔名」；無→禁「依範本」。被禁(disabled)的選項須符合範本定義 ──
const LOCK = {  // [group 禁用, index 禁用]  (filename=禁依檔名→鎖依範本; template=禁依範本→鎖依檔名)
  h1_empty: ["template", "template"], v1_empty: ["template", "template"],   // 兩軸皆無定義 → 都鎖依檔名
  h2_index: ["template", "filename"], v2_index: ["template", "filename"],   // 只有 index → index 鎖依範本
  h3_index_group: ["filename", "filename"], v3_index_group: ["filename", "filename"], // 兩軸都鎖依範本
  h4_group: ["filename", "template"], v4_group: ["filename", "template"],   // 只有 group → group 鎖依範本
};
let lockFail = 0, lockRun = 0;
for (const [name, exp] of Object.entries(LOCK)) {
  lockRun++;
  const orient = FORCE_ORIENT(name);
  const d = F.detectExcelTable(grids[name]); const ax = F.readAxisLabels(grids[name], d);
  const groupHas = (orient === "vertical" ? ax.topLabelsArr : ax.leftLabelsArr).length > 0;
  const idxHas = (orient === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr).length > 0;
  const got = [groupHas ? "filename" : "template", idxHas ? "filename" : "template"];
  try { assert.deepStrictEqual(got, exp); }
  catch (e) { lockFail++; console.log(`  ✗ 鎖定 ${name}: 期望禁 ${JSON.stringify(exp)} 實際禁 ${JSON.stringify(got)}`); }
}
console.log(`  對位鎖定：${lockRun - lockFail}/${lockRun} 範本通過`);
fail += lockFail;

// ── 寫入標籤：Index=依檔名(範本該軸無 index)時，index 標頭須寫入(與 group 對稱)；範本有 index 則不重寫 ──
function writtenLabels(gridName, orient) {
  const grid = grids[gridName]; const d = F.detectExcelTable(grid);
  const excel = { startCell: d.startCell, snCol: d.snCol, cellCols: 1, cellRows: 1, gapRows: 0, orient };
  const ax = F.readAxisLabels(grid, d);
  const idxAxis = (orient === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr);
  const label = { pattern: "{group}-{idx}", idxSort: "auto", idxOrder: [], groupSrc: "auto", idxIgnore: [], font_pt: 12 };
  if (idxAxis.length) { label.idxSort = "custom"; label.idxOrder = [...idxAxis]; }
  const cells = F.computeExcelCellsAuto(grid, excel, label, FILES, true);
  return cells.filter((c) => c.text != null).map((c) => get_column_letter(c.col) + c.row + "=" + c.text).sort();
}
const LABELS = {
  // 範本無 index → 寫入 index 標頭(group 依範本故不寫 group)
  h4_group: ["C2=1", "D2=2", "E2=3"],
  v4_group: ["B3=1", "B4=2", "B5=3"],
  // 範本有 index(2,3,4) → 不重寫；group 無定義→依檔名寫 group 標籤
  h2_index: ["B3=AAA", "B4=BBB"],
  v2_index: ["C2=AAA", "D2=BBB"],
};
let labFail = 0, labRun = 0;
for (const [name, exp] of Object.entries(LABELS)) {
  labRun++;
  const got = writtenLabels(name, FORCE_ORIENT(name));
  try { assert.deepStrictEqual(got, exp.slice().sort()); }
  catch (e) { labFail++; console.log(`  ✗ 寫入標籤 ${name}: 期望 ${JSON.stringify(exp)} 實際 ${JSON.stringify(got)}`); }
}
console.log(`  寫入標籤：${labRun - labFail}/${labRun} 範本通過`);
fail += labFail;

// ── 回歸（#24/#31）：範本 index 標籤之間有「空欄/空列」→ image 須對齊「標籤實際位置」、跳過空隙 ──
//   曾把 index 落點寫死成連續 start+ci*step（group 軸有查標籤、index 軸漏了）→ 範本有間隙就錯位。
function gridFromCells(cellList) {
  const cells = cellList.map(([r, c, text]) => ({ r, c, text }));
  const rs = cellList.map((x) => x[0]), cs = cellList.map((x) => x[1]);
  const minR = Math.min(...rs), maxR = Math.max(...rs), minC = Math.min(...cs), maxC = Math.max(...cs);
  const borders = [];
  for (let r = minR; r <= maxR; r++) for (let c = minC; c <= maxC; c++) borders.push({ r, c });
  const cols = []; for (let c = 1; c <= maxC + 3; c++) cols.push({ letter: get_column_letter(c) });
  return { cols, cells, borders };
}
function placeGap(grid, orient) {
  const d = F.detectExcelTable(grid);
  const ax = F.readAxisLabels(grid, d);
  const excel = { startCell: d.startCell, snCol: d.snCol, cellCols: 1, cellRows: 1, gapRows: 0, orient };
  const idxAxis = orient === "vertical" ? ax.leftLabelsArr : ax.topLabelsArr;
  const label = { pattern: "{group}-{idx}", idxSort: "custom", idxOrder: [...idxAxis], groupSrc: "auto", idxIgnore: [], font_pt: 12 };
  const out = {};
  for (const c of F.computeExcelCellsAuto(grid, excel, label, FILES, true)) if (c.path) out[c.path.replace(/\.[^.]+$/, "")] = get_column_letter(c.col) + c.row;
  return out;
}
const GAP = {
  // 横式：index 標籤 2@C2、(D2 空)、3@E2 → BBB-3 須到 E（跳過空欄 D），非連續的 D
  "横式·index 標籤有空欄": { grid: gridFromCells([[2,2,"Title"],[2,3,"2"],[2,5,"3"],[3,2,"BBB"],[4,2,"CCC"]]),
    orient: "horizontal", expect: { "BBB-2": "C3", "BBB-3": "E3" } },
  // 直式：index 標籤 2@B3、(B4 空)、3@B5 → BBB-3 須到第 5 列（跳過空列 4）
  "直式·index 標籤有空列": { grid: gridFromCells([[2,2,"Title"],[2,3,"BBB"],[2,4,"CCC"],[3,2,"2"],[5,2,"3"]]),
    orient: "vertical", expect: { "BBB-2": "C3", "BBB-3": "C5" } },
};
for (const [name, t] of Object.entries(GAP)) {
  const got = placeGap(t.grid, t.orient);
  try { assert.deepStrictEqual(got, t.expect); console.log(`  ✓ ${name}`); }
  catch (e) { fail++; console.log(`  ✗ ${name}\n      expected ${JSON.stringify(t.expect)}\n      got      ${JSON.stringify(got)}`); }
}

console.log(fail ? `\n${fail} case(s) FAILED` : "\nAll cases passed ✓");
process.exit(fail ? 1 : 0);
