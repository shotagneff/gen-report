/**
 * CSV 出力（BOM付きUTF-8、RFC 4180 に沿ったエスケープ）
 */

import fs from "node:fs";
import path from "node:path";

function escapeCell(value: string): string {
  const s = String(value);
  if (/[\r\n,"]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

export function toCsvRow(columns: string[]): string {
  return columns.map(escapeCell).join(",");
}

export function writeCsv(
  headers: string[],
  rows: string[][],
  outPath: string,
  bom = true,
): void {
  const dir = path.dirname(outPath);
  if (dir) fs.mkdirSync(dir, { recursive: true });
  const line1 = toCsvRow(headers);
  const rest = rows.map((row) => toCsvRow(row));
  const content = (bom ? "\uFEFF" : "") + [line1, ...rest].join("\n") + "\n";
  fs.writeFileSync(outPath, content, "utf8");
}
