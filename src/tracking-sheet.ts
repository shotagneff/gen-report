/**
 * リード管理CRM
 * レポート生成のたびに指定フォルダ内の「リード管理CRM」に1行追加する。
 * 初回実行時はシートを自動作成する。
 */

import fs from "node:fs";
import path from "node:path";
import { google } from "googleapis";
import type { drive_v3, sheets_v4 } from "googleapis";
import { datePrefix } from "./sheets-export.js";

const TRACKING_SHEET_NAME = "リード管理CRM";
const SHEET_TAB = "リスト";
const HEADERS = [
  "作成日", "会社名", "ホームページURL", "住所", "電話番号",
  "レポートURL", "フォーム営業文", "ステータス",
  "スコア", "ランク", "スコアリング日", "接触経路", "反応メモ", "推奨アクション",
];
const COL_COUNT = HEADERS.length; // 14
const STATUS_OPTIONS = ["未アプローチ", "アプローチ済み", "フォーム営業完了", "Aランク対応中", "ナーチャリング中", "3ヶ月後フォロー"];
const RANK_OPTIONS = ["A", "B", "C"];
const CONTACT_PATH_OPTIONS = ["フォーム", "テレアポ", "訪問", "メール返信", "その他"];

// 列幅: A〜N
const COL_WIDTHS = [120, 180, 220, 200, 130, 120, 300, 120, 80, 60, 120, 120, 300, 300];

// チップ風スタイル（レポートURL列）
const CHIP_BG   = { red: 0.788, green: 0.855, blue: 0.973 };
const CHIP_TEXT = { red: 0.118, green: 0.227, blue: 0.376 };
const HEADER_BG = { red: 0.118, green: 0.227, blue: 0.376 };
const WHITE     = { red: 1,     green: 1,     blue: 1     };

interface TrackingRow {
  date: string;
  companyName: string;
  siteUrl: string;
  address: string;
  phone: string;
  reportUrl: string;
  outreachMessage: string;
  status: string;
  score?: string;
  rank?: string;
  scoringDate?: string;
  contactPath?: string;
  responseNotes?: string;
  recommendedAction?: string;
}

/** 既存CRMのヘッダーを最新形式にマイグレーションする */
async function migrateHeadersIfNeeded(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetId: number,
  tabTitle: string,
): Promise<void> {
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${tabTitle}'!A1:N1`,
  });
  const currentHeaders = headerRes.data.values?.[0] ?? [];

  // 旧形式1: 7列（ステータスがG列）→ 14列
  if (currentHeaders.length === 7 && currentHeaders[6] === "ステータス") {
    // G列（ステータス）の前に列を挿入
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            insertDimension: {
              range: { sheetId, dimension: "COLUMNS", startIndex: 6, endIndex: 7 },
              inheritFromBefore: false,
            },
          },
        ],
      },
    });

    // 新G列のヘッダーを「フォーム営業文」に設定
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${tabTitle}'!G1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [["フォーム営業文"]] },
    });

    // 現在8列になったので、以下の8列→14列マイグレーションに落ちる
    currentHeaders.splice(6, 0, "フォーム営業文");
  }

  // 旧形式2: 8列（フォーム営業文G列+ステータスH列）→ 14列
  if (currentHeaders.length === 8 && currentHeaders[7] === "ステータス") {
    // I〜N列のヘッダーを追加
    const newHeaders = HEADERS.slice(8); // ["スコア", "ランク", "スコアリング日", "接触経路", "反応メモ", "推奨アクション"]
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${tabTitle}'!I1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [newHeaders] },
    });

    // スタイルとバリデーションを適用
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          // I〜N列ヘッダーのスタイル
          {
            repeatCell: {
              range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 8, endColumnIndex: COL_COUNT },
              cell: {
                userEnteredFormat: {
                  backgroundColor: HEADER_BG,
                  textFormat: { foregroundColor: WHITE, bold: true },
                  horizontalAlignment: "CENTER",
                },
              },
              fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            },
          },
          // I〜N列の列幅
          ...COL_WIDTHS.slice(8).map((pixels, i) => ({
            updateDimensionProperties: {
              range: { sheetId, dimension: "COLUMNS", startIndex: 8 + i, endIndex: 9 + i },
              properties: { pixelSize: pixels },
              fields: "pixelSize",
            },
          })),
          // ステータス列（H列）のドロップダウンを拡張ステータスで更新
          {
            setDataValidation: {
              range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 7, endColumnIndex: 8 },
              rule: {
                condition: {
                  type: "ONE_OF_LIST",
                  values: STATUS_OPTIONS.map((v) => ({ userEnteredValue: v })),
                },
                showCustomUi: true,
                strict: false,
              },
            },
          },
          // J列（ランク）にA/B/Cドロップダウン
          {
            setDataValidation: {
              range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 9, endColumnIndex: 10 },
              rule: {
                condition: {
                  type: "ONE_OF_LIST",
                  values: RANK_OPTIONS.map((v) => ({ userEnteredValue: v })),
                },
                showCustomUi: true,
                strict: false,
              },
            },
          },
          // L列（接触経路）にドロップダウン
          {
            setDataValidation: {
              range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 11, endColumnIndex: 12 },
              rule: {
                condition: {
                  type: "ONE_OF_LIST",
                  values: CONTACT_PATH_OPTIONS.map((v) => ({ userEnteredValue: v })),
                },
                showCustomUi: true,
                strict: false,
              },
            },
          },
        ],
      },
    });
  }
}

/** フォルダ内で管理シートを検索し、なければ新規作成して { spreadsheetId, sheetId, tabTitle } を返す */
async function findOrCreateTrackingSheet(
  drive: drive_v3.Drive,
  sheets: sheets_v4.Sheets,
  folderId: string,
): Promise<{ spreadsheetId: string; sheetId: number; tabTitle: string }> {
  // フォルダ内で管理シートを検索
  const listRes = await drive.files.list({
    q: `name='${TRACKING_SHEET_NAME}' and '${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: "files(id)",
    pageSize: 1,
  });

  const existing = listRes.data.files?.[0];
  if (existing?.id) {
    // 既存シートの場合は最初のタブ名とsheetIdを取得して返す
    const spreadsheetId = existing.id;
    const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
    const firstSheet = meta.data.sheets?.[0].properties;
    const sheetId = firstSheet?.sheetId ?? 0;
    const tabTitle = firstSheet?.title ?? SHEET_TAB;

    // ヘッダーの自動マイグレーション
    await migrateHeadersIfNeeded(sheets, spreadsheetId, sheetId, tabTitle);

    return { spreadsheetId, sheetId, tabTitle };
  }

  // 新規作成
  const createRes = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: TRACKING_SHEET_NAME },
      sheets: [{ properties: { title: SHEET_TAB } }],
    },
  });

  const spreadsheetId = createRes.data.spreadsheetId;
  if (!spreadsheetId) throw new Error("リード管理CRMの作成に失敗しました");

  const sheetId = createRes.data.sheets?.[0].properties?.sheetId ?? 0;

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
    range: `'${SHEET_TAB}'!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [HEADERS] },
  });

  // ヘッダーのスタイル + ドロップダウン設定
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // ヘッダー背景色・文字色・太字・中央揃え
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
            cell: {
              userEnteredFormat: {
                backgroundColor: HEADER_BG,
                textFormat: { foregroundColor: WHITE, bold: true },
                horizontalAlignment: "CENTER",
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
          },
        },
        // 列幅設定
        ...COL_WIDTHS.map((pixels, i) => ({
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
        // ヘッダー行折り返し
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        },
        // ヘッダー行を固定
        {
          updateSheetProperties: {
            properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
            fields: "gridProperties.frozenRowCount",
          },
        },
        // ステータス列（H列 = index7）にドロップダウン
        {
          setDataValidation: {
            range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 7, endColumnIndex: 8 },
            rule: {
              condition: {
                type: "ONE_OF_LIST",
                values: STATUS_OPTIONS.map((v) => ({ userEnteredValue: v })),
              },
              showCustomUi: true,
              strict: false,
            },
          },
        },
        // ランク列（J列 = index9）にA/B/Cドロップダウン
        {
          setDataValidation: {
            range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 9, endColumnIndex: 10 },
            rule: {
              condition: {
                type: "ONE_OF_LIST",
                values: RANK_OPTIONS.map((v) => ({ userEnteredValue: v })),
              },
              showCustomUi: true,
              strict: false,
            },
          },
        },
        // 接触経路列（L列 = index11）にドロップダウン
        {
          setDataValidation: {
            range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 11, endColumnIndex: 12 },
            rule: {
              condition: {
                type: "ONE_OF_LIST",
                values: CONTACT_PATH_OPTIONS.map((v) => ({ userEnteredValue: v })),
              },
              showCustomUi: true,
              strict: false,
            },
          },
        },
      ],
    },
  });

  return { spreadsheetId, sheetId, tabTitle: SHEET_TAB };
}

/** トラッキングシートの末尾に1行追加してフォーマットを適用する */
async function appendTrackingRow(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetId: number,
  tabTitle: string,
  row: TrackingRow,
): Promise<void> {
  const appendRes = await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${tabTitle}'!A:N`,
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
        row.outreachMessage,
        row.status,
        row.score ?? "",
        row.rank ?? "",
        row.scoringDate ?? "",
        row.contactPath ?? "",
        row.responseNotes ?? "",
        row.recommendedAction ?? "",
      ]],
    },
  });

  // 追加された行のインデックスを取得
  const updatedRange = appendRes.data.updates?.updatedRange ?? "";
  const match = updatedRange.match(/(\d+)(?::.*)?$/);
  const rowIndex = match ? parseInt(match[1], 10) - 1 : -1;
  if (rowIndex < 1) return; // ヘッダー行は変更しない

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // 行全体の背景色を白・文字色を黒にリセット（ヘッダーの白文字が引き継がれないよう）
        {
          repeatCell: {
            range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
            cell: {
              userEnteredFormat: {
                backgroundColor: WHITE,
                textFormat: { foregroundColor: { red: 0, green: 0, blue: 0 }, bold: false },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        },
        // レポートURL列（F列 = index5）をチップ風スタイルに
        {
          repeatCell: {
            range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 5, endColumnIndex: 6 },
            cell: {
              userEnteredFormat: {
                backgroundColor: CHIP_BG,
                textFormat: { bold: true, foregroundColor: CHIP_TEXT },
                horizontalAlignment: "CENTER",
                verticalAlignment: "MIDDLE",
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
          },
        },
      ],
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
  outreachMessage?: string;
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

  const { spreadsheetId, sheetId, tabTitle } = await findOrCreateTrackingSheet(drive, sheets, args.folderId);

  const date = datePrefix().replace(/_$/, "");

  await appendTrackingRow(sheets, spreadsheetId, sheetId, tabTitle, {
    date,
    companyName: args.companyName,
    siteUrl: args.siteUrl,
    address: args.address,
    phone: args.phone,
    reportUrl: args.reportUrl,
    outreachMessage: args.outreachMessage ?? "",
    status: "未アプローチ",
  });
}
