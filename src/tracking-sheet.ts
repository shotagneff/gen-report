/**
 * 営業管理トラッキングシート
 * レポート生成のたびに指定フォルダ内の「【AIレポート管理】営業リスト」に1行追加する。
 * 初回実行時はシートを自動作成する。
 */

import fs from "node:fs";
import path from "node:path";
import { google } from "googleapis";
import type { drive_v3, sheets_v4 } from "googleapis";
import { datePrefix } from "./sheets-export.js";

const TRACKING_SHEET_NAME = "【AIレポート管理】営業リスト";
const HEADERS = ["作成日", "会社名", "ホームページURL", "住所", "電話番号", "レポートURL", "ステータス"];

interface TrackingRow {
  date: string;
  companyName: string;
  siteUrl: string;
  address: string;
  phone: string;
  reportUrl: string;
  status: string;
}

/** フォルダ内で管理シートを検索し、なければ新規作成してspreadsheetIdを返す */
async function findOrCreateTrackingSheet(
  drive: drive_v3.Drive,
  sheets: sheets_v4.Sheets,
  folderId: string,
): Promise<string> {
  // フォルダ内で管理シートを検索
  const listRes = await drive.files.list({
    q: `name='${TRACKING_SHEET_NAME}' and '${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: "files(id)",
    pageSize: 1,
  });

  const existing = listRes.data.files?.[0];
  if (existing?.id) {
    return existing.id;
  }

  // 新規作成
  const createRes = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: TRACKING_SHEET_NAME },
      sheets: [{ properties: { title: "営業リスト" } }],
    },
  });

  const spreadsheetId = createRes.data.spreadsheetId;
  if (!spreadsheetId) throw new Error("トラッキングシートの作成に失敗しました");

  // 指定フォルダへ移動
  const fileRes = await drive.files.get({ fileId: spreadsheetId, fields: "parents" });
  const currentParents = fileRes.data.parents?.join(",") ?? "";
  await drive.files.update({
    fileId: spreadsheetId,
    addParents: folderId,
    removeParents: currentParents,
    requestBody: {},
  });

  // ヘッダー行を書き込む
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: "'営業リスト'!A1",
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [HEADERS] },
  });

  // ヘッダー行のスタイルを設定（濃紺背景・白文字・太字・列幅）
  const sheetId = createRes.data.sheets?.[0].properties?.sheetId ?? 0;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // ヘッダー背景色・文字色・太字
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: HEADERS.length },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.118, green: 0.227, blue: 0.376 },
                textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true },
                horizontalAlignment: "CENTER",
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
          },
        },
        // 列幅設定
        ...([200, 180, 220, 200, 130, 280, 120] as number[]).map((pixels, i) => ({
          updateDimensionProperties: {
            range: { sheetId, dimension: "COLUMNS", startIndex: i, endIndex: i + 1 },
            properties: { pixelSize: pixels },
            fields: "pixelSize",
          },
        })),
        // ヘッダー行の高さ
        {
          updateDimensionProperties: {
            range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: 1 },
            properties: { pixelSize: 40 },
            fields: "pixelSize",
          },
        },
        // 行の折り返し設定（ヘッダー）
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        },
        // ウィンドウ枠の固定（ヘッダー行を固定）
        {
          updateSheetProperties: {
            properties: {
              sheetId,
              gridProperties: { frozenRowCount: 1 },
            },
            fields: "gridProperties.frozenRowCount",
          },
        },
      ],
    },
  });

  return spreadsheetId;
}

/** トラッキングシートの末尾に1行追加する */
async function appendTrackingRow(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  row: TrackingRow,
): Promise<void> {
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: "'営業リスト'!A:G",
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[
        row.date,
        row.companyName,
        row.siteUrl,
        row.address,
        row.phone,
        row.reportUrl,
        row.status,
      ]],
    },
  });
}

/** スクリプトから呼び出す窓口。認証・検索・追記を一括処理する。 */
export async function updateTracking(args: {
  companyName: string;
  siteUrl: string;
  address: string;
  phone: string;
  reportUrl: string;
  folderId: string;
  credentialsPath: string;
}): Promise<void> {
  const keyPath = path.resolve(args.credentialsPath);
  if (!fs.existsSync(keyPath)) {
    throw new Error(`認証ファイルが見つかりません: ${keyPath}`);
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: keyPath,
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

  const spreadsheetId = await findOrCreateTrackingSheet(drive, sheets, args.folderId);

  // 末尾の "_" を除いた日付文字列（例: "2026年02月21日"）
  const date = datePrefix().replace(/_$/, "");

  await appendTrackingRow(sheets, spreadsheetId, {
    date,
    companyName: args.companyName,
    siteUrl: args.siteUrl,
    address: args.address,
    phone: args.phone,
    reportUrl: args.reportUrl,
    status: "未アプローチ",
  });
}
