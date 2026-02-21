#!/usr/bin/env npx tsx
/**
 * リード管理CRMのスコアリング列（I〜N）を更新するスクリプト。
 * Usage: npx tsx scripts/update-crm-score.ts --company-name "株式会社〇〇" --score 72 --rank A --contact-type "フォーム" --memo "反応良好" --action "24h以内に電話"
 */

import "dotenv/config";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { google } from "googleapis";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");

const TRACKING_SHEET_NAME = "リード管理CRM";

// ランクに応じたステータス自動更新マッピング
const RANK_STATUS_MAP: Record<string, string> = {
  A: "Aランク対応中",
  B: "ナーチャリング中",
  C: "3ヶ月後フォロー",
};

function parseArgs(argv: string[]): {
  companyName: string;
  score: string;
  rank: string;
  contactType: string;
  memo: string;
  action: string;
  updateStatus: boolean;
} {
  let companyName = "";
  let score = "";
  let rank = "";
  let contactType = "";
  let memo = "";
  let action = "";
  let updateStatus = true;

  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--company-name" && argv[i + 1]) companyName = argv[++i];
    else if (argv[i] === "--score" && argv[i + 1]) score = argv[++i];
    else if (argv[i] === "--rank" && argv[i + 1]) rank = argv[++i];
    else if (argv[i] === "--contact-type" && argv[i + 1]) contactType = argv[++i];
    else if (argv[i] === "--memo" && argv[i + 1]) memo = argv[++i];
    else if (argv[i] === "--action" && argv[i + 1]) action = argv[++i];
    else if (argv[i] === "--no-status-update") updateStatus = false;
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/update-crm-score.ts --company-name <name> --score <n> --rank <A|B|C> [options]

  リード管理CRMのスコアリング列（I〜N）を更新します。

  --company-name NAME     会社名（部分一致で検索）
  --score N               スコア値（0〜100）
  --rank RANK             ランク（A, B, C）
  --contact-type TYPE     接触経路（フォーム/テレアポ/訪問/メール返信/その他）
  --memo TEXT             反応メモ
  --action TEXT           推奨アクション
  --no-status-update      ランクに応じたステータス自動更新をスキップ
`);
      process.exit(0);
    }
  }
  return { companyName, score, rank, contactType, memo, action, updateStatus };
}

async function main(): Promise<void> {
  const { companyName, score, rank, contactType, memo, action, updateStatus } = parseArgs(process.argv.slice(2));

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

  // 会社名で検索
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
  const rowNum = matchRow + 1; // 1-indexed

  // スコアリング日を自動設定
  const now = new Date();
  const scoringDate = `${String(now.getFullYear()).slice(2)}_${String(now.getMonth() + 1).padStart(2, "0")}_${String(now.getDate()).padStart(2, "0")}`;

  // I〜N列を一括更新
  const updateValues = [score, rank, scoringDate, contactType, memo, action];
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${tabTitle}'!I${rowNum}:N${rowNum}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [updateValues] },
  });

  console.error(`✅ 「${matchedName}」のスコアリングデータを更新しました（行: ${rowNum}）`);
  console.error(`   スコア: ${score} / ランク: ${rank} / 接触経路: ${contactType}`);

  // ランクに応じてステータスを自動更新
  if (updateStatus && rank && RANK_STATUS_MAP[rank]) {
    const statusRange = `'${tabTitle}'!H${rowNum}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: statusRange,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[RANK_STATUS_MAP[rank]]] },
    });
    console.error(`   ステータスを「${RANK_STATUS_MAP[rank]}」に更新しました`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
