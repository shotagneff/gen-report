#!/usr/bin/env npx tsx
/**
 * アクティビティ履歴の記録・一覧表示スクリプト。
 * Usage:
 *   npx tsx scripts/log-crm-activity.ts --company "株式会社〇〇" --type "電話" --content "決裁者と面談" --result "提案書送付依頼"
 *   npx tsx scripts/log-crm-activity.ts --company "株式会社〇〇" --list
 */

import "dotenv/config";
import { getCRMConnection, formatDateTimeStamp } from "../src/crm-common.js";

const TAB_ACTIVITIES = "アクティビティ";

function parseArgs(argv: string[]): {
  companyName: string;
  list: boolean;
  type: string;
  contactName: string;
  content: string;
  result: string;
} {
  let companyName = "";
  let list = false;
  let type = "";
  let contactName = "";
  let content = "";
  let result = "";

  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--company" && argv[i + 1]) companyName = argv[++i];
    else if (argv[i] === "--list") list = true;
    else if (argv[i] === "--type" && argv[i + 1]) type = argv[++i];
    else if (argv[i] === "--contact" && argv[i + 1]) contactName = argv[++i];
    else if (argv[i] === "--content" && argv[i + 1]) content = argv[++i];
    else if (argv[i] === "--result" && argv[i + 1]) result = argv[++i];
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/log-crm-activity.ts [options]

  アクティビティ履歴の記録・一覧表示。

  --company NAME      会社名（必須）
  --list              指定企業のアクティビティ一覧を表示
  --type TYPE         種別（メール送信/メール返信受信/電話/訪問/スコアリング更新/ステータス変更/フォーム営業/その他）
  --contact NAME      担当者名
  --content TEXT      内容
  --result TEXT       結果
`);
      process.exit(0);
    }
  }
  return { companyName, list, type, contactName, content, result };
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));

  if (!args.companyName) {
    console.error("Error: --company is required");
    process.exit(2);
  }

  const conn = await getCRMConnection();
  const activityTab = conn.tabs.get(TAB_ACTIVITIES);

  if (!activityTab) {
    console.error("Error: アクティビティタブが見つかりません。gen-reportを一度実行してCRMタブを作成してください。");
    process.exit(1);
  }

  if (args.list) {
    // 一覧表示
    const dataRes = await conn.sheets.spreadsheets.values.get({
      spreadsheetId: conn.spreadsheetId,
      range: `'${TAB_ACTIVITIES}'!A:G`,
    });
    const rows = dataRes.data.values ?? [];
    const filtered = rows.filter((row, i) => {
      if (i === 0) return false;
      const cell = String(row[1] ?? "");
      return cell.includes(args.companyName) || args.companyName.includes(cell);
    });

    if (filtered.length === 0) {
      console.error(`「${args.companyName}」のアクティビティが見つかりません`);
      process.exit(1);
    }

    const result = filtered.map((row) => ({
      datetime: row[0] ?? "",
      company: row[1] ?? "",
      type: row[2] ?? "",
      contact: row[3] ?? "",
      content: row[4] ?? "",
      result: row[5] ?? "",
      recorder: row[6] ?? "",
    }));
    console.log(JSON.stringify(result, null, 2));
  } else {
    // 記録
    if (!args.type || !args.content) {
      console.error("Error: --type と --content は必須です");
      process.exit(2);
    }

    const timestamp = formatDateTimeStamp();
    await conn.sheets.spreadsheets.values.append({
      spreadsheetId: conn.spreadsheetId,
      range: `'${TAB_ACTIVITIES}'!A:G`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          timestamp,
          args.companyName,
          args.type,
          args.contactName,
          args.content,
          args.result,
          "手動",
        ]],
      },
    });

    console.error(`✅ アクティビティを記録しました: ${args.companyName} / ${args.type}`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
