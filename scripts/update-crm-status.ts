#!/usr/bin/env npx tsx
/**
 * リード管理CRMのステータス列を更新するスクリプト。
 * Usage: npx tsx scripts/update-crm-status.ts --company-name "株式会社〇〇" --status "アプローチ済み"
 */

import "dotenv/config";
import { getCRMConnection, findCompanyRow } from "../src/crm-common.js";

function parseArgs(argv: string[]): { companyName: string; status: string; skipStatus: boolean; outreachMessage: string } {
  let companyName = "";
  let status = "アプローチ済み";
  let skipStatus = true; // デフォルトはステータス更新しない
  let outreachMessage = "";
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--company-name" && argv[i + 1]) companyName = argv[++i];
    else if (argv[i] === "--status" && argv[i + 1]) { status = argv[++i]; skipStatus = false; }
    else if (argv[i] === "--outreach-message" && argv[i + 1]) outreachMessage = argv[++i];
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/update-crm-status.ts --company-name <name> [--status <status>] [--outreach-message <message>]

  リード管理CRMの指定会社のステータスとフォーム営業文を更新します。

  --company-name NAME        会社名（部分一致で検索）
  --status STATUS            設定するステータス（デフォルト: アプローチ済み）
  --outreach-message MSG     フォーム営業文（G列に書き込み）
`);
      process.exit(0);
    }
  }
  return { companyName, status, skipStatus, outreachMessage };
}

async function main(): Promise<void> {
  const { companyName, status, skipStatus, outreachMessage } = parseArgs(process.argv.slice(2));

  if (!companyName) {
    console.error("Error: --company-name is required");
    process.exit(2);
  }

  const conn = await getCRMConnection();
  const tabTitle = conn.tabs.entries().next().value?.[1].title ?? "リスト";

  const match = await findCompanyRow(conn.sheets, conn.spreadsheetId, tabTitle, companyName);
  if (!match) {
    console.error(`Error: 「${companyName}」に一致する会社が見つかりません`);
    process.exit(1);
  }

  const matchedName = match.rowData[1];
  const rowNum = match.rowIndex;

  // ステータスを更新（--status が明示指定された場合のみ）
  if (!skipStatus) {
    await conn.sheets.spreadsheets.values.update({
      spreadsheetId: conn.spreadsheetId,
      range: `'${tabTitle}'!H${rowNum}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[status]] },
    });
    console.error(`✅ 「${matchedName}」のステータスを「${status}」に更新しました（行: ${rowNum}）`);

    // アクティビティ記録
    try {
      const { appendActivityRow } = await import("../src/tracking-sheet.js");
      await appendActivityRow(conn.sheets, conn.spreadsheetId, {
        companyName: matchedName,
        activityType: "ステータス変更",
        content: `ステータスを「${status}」に変更`,
        recorder: "手動",
      });
    } catch (_) {
      // アクティビティ記録失敗は無視
    }
  }

  // フォーム営業文を更新（指定された場合）
  if (outreachMessage) {
    await conn.sheets.spreadsheets.values.update({
      spreadsheetId: conn.spreadsheetId,
      range: `'${tabTitle}'!G${rowNum}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[outreachMessage]] },
    });
    console.error(`✅ フォーム営業文をG列に書き込みました`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
