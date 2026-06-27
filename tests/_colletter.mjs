// 1-based 欄號 → 字母（A,B,...,Z,AA,...）
export function get_column_letter(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
