#!/usr/bin/env npx tsx
/**
 * CRMから指定行を削除するスクリプト。
 * Usage:
 *   npx tsx scripts/delete-crm-rows.ts --rows 3,4,5,6,7,8,9,10,11,12,33
 *   npx tsx scripts/delete-crm-rows.ts --company-name "SEEKAD" --dry-run
 *   npx tsx scripts/delete-crm-rows.ts --company-name "SEEKAD"
 */

import "dotenv/config";
import { getCRMConnection } from "../src/crm-common.js";

function parseArgs(argv: string[]): { rows: number[]; companyName: string; dryRun: boolean } {
  let rows: number[] = [];
  let companyName = "";
  let dryRun = false;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--rows" && argv[i + 1]) {
      rows = argv[++i].split(",").map(Number).filter(n => n > 1);
    } else if (argv[i] === "--company-name" && argv[i + 1]) {
      companyName = argv[++i];
    } else if (argv[i] === "--dry-run") {
      dryRun = true;
    } else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/delete-crm-rows.ts [options]

  --rows 3,4,5          削除する行番号（カンマ区切り、1-indexed、ヘッダー行1は削除不可）
  --company-name NAME   会社名で検索して該当行を全削除
  --dry-run             実際には削除せず対象行を表示のみ
`);
      process.exit(0);
    }
  }
  return { rows, companyName, dryRun };
}

async function main(): Promise<void> {
  const { rows, companyName, dryRun } = parseArgs(process.argv.slice(2));

  const conn = await getCRMConnection();
  const tabEntry = conn.tabs.entries().next().value;
  const tabTitle = tabEntry?.[1].title ?? "リスト";
  const sheetId = tabEntry?.[1].sheetId ?? 0;

  let targetRows = rows;

  // 会社名指定の場合、全マッチ行を検索
  if (companyName && targetRows.length === 0) {
    const dataRes = await conn.sheets.spreadsheets.values.get({
      spreadsheetId: conn.spreadsheetId,
      range: `'${tabTitle}'!A:V`,
    });
    const allRows = dataRes.data.values ?? [];
    for (let i = 1; i < allRows.length; i++) {
      const cellValue = String(allRows[i][1] ?? "");
      if (cellValue.includes(companyName) || companyName.includes(cellValue)) {
        targetRows.push(i + 1);
      }
    }
  }

  if (targetRows.length === 0) {
    console.error("削除対象の行が見つかりません");
    process.exit(1);
  }

  // 行番号を降順にソート（下から削除しないと行番号がずれる）
  targetRows.sort((a, b) => b - a);

  if (dryRun) {
    console.error(`[dry-run] 削除対象: ${targetRows.length}行`);
    console.error(`行番号: ${targetRows.sort((a, b) => a - b).join(", ")}`);
    return;
  }

  // batchUpdateで一括削除（降順で実行）
  const requests = targetRows.map(row => ({
    deleteDimension: {
      range: {
        sheetId,
        dimension: "ROWS" as const,
        startIndex: row - 1, // 0-indexed
        endIndex: row,       // exclusive
      },
    },
  }));

  await conn.sheets.spreadsheets.batchUpdate({
    spreadsheetId: conn.spreadsheetId,
    requestBody: { requests },
  });

  console.error(`✅ ${targetRows.length}行を削除しました（行番号: ${targetRows.sort((a, b) => a - b).join(", ")}）`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
