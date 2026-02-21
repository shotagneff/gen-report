#!/usr/bin/env npx tsx
/**
 * リード管理CRMのステータス列を更新するスクリプト。
 * Usage: npx tsx scripts/update-crm-status.ts --company-name "株式会社〇〇" --status "アプローチ済み"
 */

import "dotenv/config";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { google } from "googleapis";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");

const TRACKING_SHEET_NAME = "リード管理CRM";

function loadDotenv(): void {
  const envPath = path.join(repoRoot, ".env");
  if (!fs.existsSync(envPath)) return;
  const content = fs.readFileSync(envPath, "utf8");
  for (const line of content.split(/\r?\n/)) {
    const t = line.trim();
    if (!t || t.startsWith("#")) continue;
    const eq = t.indexOf("=");
    if (eq === -1) continue;
    const key = t.slice(0, eq).trim();
    let val = t.slice(eq + 1).trim();
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1);
    }
    if (!process.env[key]) process.env[key] = val;
  }
}

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
  loadDotenv();
  const { companyName, status, skipStatus, outreachMessage } = parseArgs(process.argv.slice(2));

  if (!companyName) {
    console.error("Error: --company-name is required");
    process.exit(2);
  }

  const keyPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  if (!keyPath || !folderId) {
    console.error("Error: GOOGLE_APPLICATION_CREDENTIALS と GOOGLE_DRIVE_FOLDER_ID を .env に設定してください");
    process.exit(2);
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: path.resolve(keyPath),
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ],
    clientOptions: {
      subject: process.env.GOOGLE_IMPERSONATE_USER,
    },
  });

  const drive = google.drive({ version: "v3", auth });
  const sheets = google.sheets({ version: "v4", auth });

  // リード管理CRMを検索
  const listRes = await drive.files.list({
    q: `name='${TRACKING_SHEET_NAME}' and '${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: "files(id)",
    pageSize: 1,
  });

  const file = listRes.data.files?.[0];
  if (!file?.id) {
    console.error("Error: リード管理CRMが見つかりません");
    process.exit(1);
  }

  const spreadsheetId = file.id;

  // タブ名を取得
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
  const tabTitle = meta.data.sheets?.[0].properties?.title ?? "リスト";

  // 全データを読み取り
  const dataRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${tabTitle}'!A:N`,
  });

  const rows = dataRes.data.values ?? [];
  if (rows.length <= 1) {
    console.error("Error: CRMにデータがありません");
    process.exit(1);
  }

  // 会社名列（B列 = index 1）で部分一致検索
  let matchRow = -1;
  for (let i = 1; i < rows.length; i++) {
    const cellValue = rows[i][1] ?? "";
    if (cellValue.includes(companyName) || companyName.includes(cellValue)) {
      matchRow = i;
      break;
    }
  }

  if (matchRow < 0) {
    console.error(`Error: 「${companyName}」に一致する会社が見つかりません`);
    process.exit(1);
  }

  const matchedName = rows[matchRow][1];
  const statusRange = `'${tabTitle}'!H${matchRow + 1}`;

  // ステータスを更新（--status が明示指定された場合のみ）
  if (!skipStatus) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: statusRange,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[status]] },
    });
    console.error(`✅ 「${matchedName}」のステータスを「${status}」に更新しました（行: ${matchRow + 1}）`);
  }

  // フォーム営業文を更新（指定された場合）
  if (outreachMessage) {
    const outreachRange = `'${tabTitle}'!G${matchRow + 1}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: outreachRange,
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
