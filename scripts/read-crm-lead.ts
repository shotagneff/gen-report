#!/usr/bin/env npx tsx
/**
 * リード管理CRMからリードデータを読み取るスクリプト。
 * Usage:
 *   npx tsx scripts/read-crm-lead.ts --company-name "株式会社〇〇"
 *   npx tsx scripts/read-crm-lead.ts --all
 *   npx tsx scripts/read-crm-lead.ts --unscored
 */

import "dotenv/config";
import { getCRMConnection } from "../src/crm-common.js";

function parseArgs(argv: string[]): { companyName: string; all: boolean; unscored: boolean } {
  let companyName = "";
  let all = false;
  let unscored = false;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--company-name" && argv[i + 1]) companyName = argv[++i];
    else if (argv[i] === "--all") all = true;
    else if (argv[i] === "--unscored") unscored = true;
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/read-crm-lead.ts [options]

  リード管理CRMからリードデータを読み取り、JSON形式で標準出力に出力します。

  --company-name NAME   指定企業のデータを取得（部分一致）
  --all                 全リードを取得
  --unscored            スコア未記入（I列空）のリードのみ取得
`);
      process.exit(0);
    }
  }
  return { companyName, all, unscored };
}

interface CRMLead {
  rowIndex: number;
  date: string;
  companyName: string;
  siteUrl: string;
  address: string;
  phone: string;
  reportUrl: string;
  outreachMessage: string;
  status: string;
  score: string;
  rank: string;
  scoringDate: string;
  contactPath: string;
  responseNotes: string;
  recommendedAction: string;
  pipeline: string;
  dealAmount: string;
  winProbability: string;
  expectedCloseDate: string;
  contactName: string;
  contactEmail: string;
  contactDepartment: string;
  lastContactDate: string;
}

function rowToLead(row: string[], rowIndex: number): CRMLead {
  return {
    rowIndex,
    date: row[0] ?? "",
    companyName: row[1] ?? "",
    siteUrl: row[2] ?? "",
    address: row[3] ?? "",
    phone: row[4] ?? "",
    reportUrl: row[5] ?? "",
    outreachMessage: row[6] ?? "",
    status: row[7] ?? "",
    score: row[8] ?? "",
    rank: row[9] ?? "",
    scoringDate: row[10] ?? "",
    contactPath: row[11] ?? "",
    responseNotes: row[12] ?? "",
    recommendedAction: row[13] ?? "",
    pipeline: row[14] ?? "",
    dealAmount: row[15] ?? "",
    winProbability: row[16] ?? "",
    expectedCloseDate: row[17] ?? "",
    contactName: row[18] ?? "",
    contactEmail: row[19] ?? "",
    contactDepartment: row[20] ?? "",
    lastContactDate: row[21] ?? "",
  };
}

async function main(): Promise<void> {
  const { companyName, all, unscored } = parseArgs(process.argv.slice(2));

  if (!companyName && !all && !unscored) {
    console.error("Error: --company-name, --all, または --unscored のいずれかを指定してください");
    process.exit(2);
  }

  const conn = await getCRMConnection();
  const tabTitle = conn.tabs.entries().next().value?.[1].title ?? "リスト";

  const dataRes = await conn.sheets.spreadsheets.values.get({
    spreadsheetId: conn.spreadsheetId,
    range: `'${tabTitle}'!A:V`,
  });

  const rows = dataRes.data.values ?? [];
  if (rows.length <= 1) {
    console.error("Error: CRMにデータがありません");
    process.exit(1);
  }

  const leads: CRMLead[] = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[1]) continue;

    if (companyName) {
      const cellValue = row[1] ?? "";
      if (cellValue.includes(companyName) || companyName.includes(cellValue)) {
        leads.push(rowToLead(row, i + 1));
      }
    } else if (unscored) {
      if (!row[8]) {
        leads.push(rowToLead(row, i + 1));
      }
    } else {
      leads.push(rowToLead(row, i + 1));
    }
  }

  if (leads.length === 0) {
    console.error(companyName ? `Error: 「${companyName}」に一致するリードが見つかりません` : "Error: 条件に一致するリードが見つかりません");
    process.exit(1);
  }

  if (companyName && leads.length === 1) {
    console.log(JSON.stringify(leads[0], null, 2));
  } else {
    console.log(JSON.stringify(leads, null, 2));
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
