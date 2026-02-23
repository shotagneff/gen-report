#!/usr/bin/env npx tsx
/**
 * ナーチャリング5ステップメールをCRMのナーチャリングタブに一括登録するスクリプト。
 * Usage:
 *   npx tsx scripts/register-nurture.ts --company-name "株式会社〇〇" --email "tanaka@example.com" --nurture-file "data/out/〇〇_nurture.json"
 */

import "dotenv/config";
import fs from "node:fs";
import path from "node:path";
import { getCRMConnection, findCompanyRow, formatDateYYMMDD } from "../src/crm-common.js";
import { appendActivityRow } from "../src/tracking-sheet.js";

interface NurtureStep {
  step: number;
  subject: string;
  body: string;
  scheduledDate: string; // YY_MM_DD
}

interface NurtureData {
  company: string;
  steps: NurtureStep[];
}

function parseArgs(argv: string[]): { companyName: string; email: string; nurtureFile: string } {
  let companyName = "";
  let email = "";
  let nurtureFile = "";

  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--company-name" && argv[i + 1]) companyName = argv[++i];
    else if (argv[i] === "--email" && argv[i + 1]) email = argv[++i];
    else if (argv[i] === "--nurture-file" && argv[i + 1]) nurtureFile = argv[++i];
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/register-nurture.ts [options]

  ナーチャリング5ステップメールをCRMに登録します。

  --company-name NAME     会社名
  --email EMAIL           送信先メールアドレス
  --nurture-file PATH     ナーチャリングメールJSONファイルパス
`);
      process.exit(0);
    }
  }
  return { companyName, email, nurtureFile };
}

async function main(): Promise<void> {
  const { companyName, email, nurtureFile } = parseArgs(process.argv.slice(2));

  if (!companyName || !email || !nurtureFile) {
    console.error("Error: --company-name, --email, --nurture-file はすべて必須です");
    process.exit(2);
  }

  // JSONファイル読み込み
  const filePath = path.resolve(nurtureFile);
  if (!fs.existsSync(filePath)) {
    console.error(`Error: ファイルが見つかりません: ${filePath}`);
    process.exit(1);
  }

  const data: NurtureData = JSON.parse(fs.readFileSync(filePath, "utf-8"));
  if (!data.steps || data.steps.length === 0) {
    console.error("Error: stepsが空です");
    process.exit(1);
  }

  const conn = await getCRMConnection();

  // ナーチャリングタブを検索
  const nurtureTab = conn.tabs.get("ナーチャリング");
  if (!nurtureTab) {
    console.error("Error: ナーチャリングタブが見つかりません。先にgen-reportを実行してタブを自動作成してください。");
    process.exit(1);
  }

  // 既存登録チェック（同じ会社名の行があれば警告）
  const existingRes = await conn.sheets.spreadsheets.values.get({
    spreadsheetId: conn.spreadsheetId,
    range: `'ナーチャリング'!A:A`,
  });
  const existingRows = existingRes.data.values ?? [];
  const alreadyRegistered = existingRows.some((row, i) =>
    i > 0 && String(row[0] ?? "").includes(companyName)
  );
  if (alreadyRegistered) {
    console.error(`⚠️ 「${companyName}」は既にナーチャリングタブに登録されています。重複登録します。`);
  }

  // 5ステップを一括追加
  const rows = data.steps
    .sort((a, b) => a.step - b.step)
    .map(s => [
      companyName,       // A: 会社名
      email,             // B: 担当者メール
      String(s.step),    // C: Step
      s.subject,         // D: 件名
      s.body,            // E: 本文
      s.scheduledDate,   // F: 送信予定日
      "",                // G: 送信日時（空=未送信）
      "",                // H: 開封
      "0",               // I: クリック回数
    ]);

  await conn.sheets.spreadsheets.values.append({
    spreadsheetId: conn.spreadsheetId,
    range: `'ナーチャリング'!A:I`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows },
  });

  console.error(`✅ 「${companyName}」のナーチャリング ${data.steps.length}ステップを登録しました`);
  console.error(`   送信先: ${email}`);
  console.error(`   送信予定: ${data.steps.map(s => `Step${s.step}=${s.scheduledDate}`).join(", ")}`);

  // リストタブのステータスを「ナーチャリング中」に更新
  const tabTitle = conn.tabs.entries().next().value?.[1].title ?? "リスト";
  const match = await findCompanyRow(conn.sheets, conn.spreadsheetId, tabTitle, companyName);
  if (match) {
    await conn.sheets.spreadsheets.values.update({
      spreadsheetId: conn.spreadsheetId,
      range: `'${tabTitle}'!H${match.rowIndex}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [["ナーチャリング中"]] },
    });
    console.error(`   ステータスを「ナーチャリング中」に更新しました`);
  }

  // アクティビティ記録
  try {
    await appendActivityRow(conn.sheets, conn.spreadsheetId, {
      companyName,
      activityType: "メール送信",
      content: `ナーチャリング ${data.steps.length}ステップ登録（送信先: ${email}）`,
      result: `Step1送信予定: ${data.steps[0]?.scheduledDate ?? "不明"}`,
      recorder: "自動",
    });
  } catch (_) {
    // アクティビティ記録失敗は無視
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
