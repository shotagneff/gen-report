#!/usr/bin/env npx tsx
/**
 * パイプライン/ディール情報の更新スクリプト。
 * Usage:
 *   npx tsx scripts/update-crm-pipeline.ts --company "株式会社〇〇" --stage "商談" --deal-amount 500000 --win-prob 60 --expected-close "2026-06-01"
 */

import "dotenv/config";
import { getCRMConnection, findCompanyRow, formatDateYYMMDD } from "../src/crm-common.js";

function parseArgs(argv: string[]): {
  companyName: string;
  stage: string;
  dealAmount: string;
  winProb: string;
  expectedClose: string;
} {
  let companyName = "", stage = "", dealAmount = "", winProb = "", expectedClose = "";

  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--company" && argv[i + 1]) companyName = argv[++i];
    else if (argv[i] === "--stage" && argv[i + 1]) stage = argv[++i];
    else if (argv[i] === "--deal-amount" && argv[i + 1]) dealAmount = argv[++i];
    else if (argv[i] === "--win-prob" && argv[i + 1]) winProb = argv[++i];
    else if (argv[i] === "--expected-close" && argv[i + 1]) expectedClose = argv[++i];
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/update-crm-pipeline.ts --company <name> [options]

  リード管理CRMのパイプライン/ディール情報（O〜R列）を更新します。

  --company NAME           会社名（部分一致で検索、必須）
  --stage STAGE            パイプラインステージ（リード/アプローチ中/商談/提案/交渉/受注/失注）
  --deal-amount N          ディール金額
  --win-prob N             受注確度（0-100%）
  --expected-close DATE    予想受注日（YYYY-MM-DD形式）
`);
      process.exit(0);
    }
  }
  return { companyName, stage, dealAmount, winProb, expectedClose };
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));

  if (!args.companyName) {
    console.error("Error: --company is required");
    process.exit(2);
  }

  const conn = await getCRMConnection();
  const listTab = conn.tabs.entries().next().value;
  const tabTitle = listTab ? listTab[1].title : "リスト";

  const match = await findCompanyRow(conn.sheets, conn.spreadsheetId, tabTitle, args.companyName);
  if (!match) {
    console.error(`Error: 「${args.companyName}」に一致する会社が見つかりません`);
    process.exit(1);
  }

  const rowNum = match.rowIndex;
  const matchedName = match.rowData[1];

  // O〜R列を一括更新（空文字のフィールドは既存値を維持するため、個別に更新）
  const updates: Array<{ range: string; values: string[][] }> = [];

  if (args.stage) {
    updates.push({ range: `'${tabTitle}'!O${rowNum}`, values: [[args.stage]] });
  }
  if (args.dealAmount) {
    updates.push({ range: `'${tabTitle}'!P${rowNum}`, values: [[args.dealAmount]] });
  }
  if (args.winProb) {
    updates.push({ range: `'${tabTitle}'!Q${rowNum}`, values: [[args.winProb]] });
  }
  if (args.expectedClose) {
    updates.push({ range: `'${tabTitle}'!R${rowNum}`, values: [[args.expectedClose]] });
  }

  if (updates.length === 0) {
    console.error("Error: 更新するフィールドが指定されていません（--stage, --deal-amount, --win-prob, --expected-close のいずれかを指定）");
    process.exit(2);
  }

  await conn.sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: conn.spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: updates,
    },
  });

  console.error(`✅ 「${matchedName}」のパイプライン情報を更新しました（行: ${rowNum}）`);
  if (args.stage) console.error(`   パイプライン: ${args.stage}`);
  if (args.dealAmount) console.error(`   ディール金額: ${args.dealAmount}`);
  if (args.winProb) console.error(`   受注確度: ${args.winProb}%`);
  if (args.expectedClose) console.error(`   予想受注日: ${args.expectedClose}`);

  // アクティビティ記録
  try {
    const { appendActivityRow } = await import("../src/tracking-sheet.js");
    await appendActivityRow(conn.sheets, conn.spreadsheetId, {
      companyName: matchedName,
      activityType: "ステータス変更",
      content: `パイプライン更新: ${[args.stage, args.dealAmount ? `金額:${args.dealAmount}` : "", args.winProb ? `確度:${args.winProb}%` : ""].filter(Boolean).join(" / ")}`,
      recorder: "手動",
    });
  } catch (_) {
    // アクティビティ記録失敗は無視
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
