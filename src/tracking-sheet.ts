/**
 * リード管理CRM
 * レポート生成のたびに指定フォルダ内の「リード管理CRM」に1行追加する。
 * 初回実行時はシートを自動作成する。
 * HubSpot/Salesforce風の5タブ構成（リスト・コンタクト・アクティビティ・タスク・ダッシュボード）。
 */

import fs from "node:fs";
import path from "node:path";
import { google } from "googleapis";
import type { drive_v3, sheets_v4 } from "googleapis";
import { datePrefix } from "./sheets-export.js";

const TRACKING_SHEET_NAME = "リード管理CRM";
const SHEET_TAB = "リスト";
const TAB_CONTACTS = "コンタクト";
const TAB_ACTIVITIES = "アクティビティ";
const TAB_TASKS = "タスク";
const TAB_DASHBOARD = "ダッシュボード";

// ==================== リストタブ定義 ====================

const HEADERS = [
  "作成日", "会社名", "ホームページURL", "住所", "電話番号",
  "レポートURL", "フォーム営業文", "ステータス",
  "スコア", "ランク", "スコアリング日", "接触経路", "反応メモ", "推奨アクション",
  "パイプライン", "ディール金額", "受注確度(%)", "予想受注日",
  "担当者名", "担当者メール", "担当者部署", "最終接触日",
];
const COL_COUNT = HEADERS.length; // 22

const STATUS_OPTIONS = ["未アプローチ", "アプローチ済み", "フォーム営業完了", "Aランク対応中", "ナーチャリング中", "3ヶ月後フォロー"];
const RANK_OPTIONS = ["A", "B", "C"];
const CONTACT_PATH_OPTIONS = ["フォーム", "テレアポ", "訪問", "メール返信", "その他"];
const PIPELINE_STAGES = ["リード", "アプローチ中", "商談", "提案", "交渉", "受注", "失注"];

// 列幅: A〜V
const COL_WIDTHS = [120, 180, 220, 200, 130, 120, 300, 120, 80, 60, 120, 120, 300, 300, 120, 100, 80, 120, 120, 180, 120, 120];

// ==================== コンタクトタブ定義 ====================

const CONTACT_HEADERS = ["コンタクトID", "会社名", "担当者名", "部署・役職", "メールアドレス", "電話番号", "キーマン", "メモ"];
const CONTACT_COL_WIDTHS = [100, 180, 120, 150, 200, 140, 100, 300];
const KEYMAN_OPTIONS = ["決裁者", "窓口", "技術担当", "その他"];

// ==================== アクティビティタブ定義 ====================

const ACTIVITY_HEADERS = ["日時", "会社名", "種別", "担当者名", "内容", "結果", "記録者"];
const ACTIVITY_COL_WIDTHS = [140, 180, 130, 120, 400, 300, 80];
const ACTIVITY_TYPES = ["メール送信", "メール返信受信", "電話", "訪問", "スコアリング更新", "ステータス変更", "フォーム営業", "その他"];

// ==================== タスクタブ定義 ====================

const TASK_HEADERS = ["タスクID", "会社名", "タスク内容", "期限", "優先度", "ステータス", "作成日", "完了日"];
const TASK_COL_WIDTHS = [100, 180, 300, 120, 80, 100, 120, 120];
const TASK_PRIORITY = ["高", "中", "低"];
const TASK_STATUS_OPTIONS = ["未着手", "進行中", "完了"];

// ==================== スタイル定数 ====================

const CHIP_BG   = { red: 0.788, green: 0.855, blue: 0.973 };
const CHIP_TEXT = { red: 0.118, green: 0.227, blue: 0.376 };
const HEADER_BG = { red: 0.118, green: 0.227, blue: 0.376 };
const WHITE     = { red: 1,     green: 1,     blue: 1     };
const BLACK     = { red: 0,     green: 0,     blue: 0     };

// 条件付き書式の色定義
const COLOR_GREEN  = { red: 0.718, green: 0.882, blue: 0.804 }; // #B7E1CD
const COLOR_YELLOW = { red: 1.0,   green: 0.949, blue: 0.800 }; // #FFF2CC
const COLOR_RED    = { red: 0.957, green: 0.780, blue: 0.765 }; // #F4C7C3
const COLOR_BLUE   = { red: 0.792, green: 0.855, blue: 0.969 }; // #CADAF7
const COLOR_GREY   = { red: 0.816, green: 0.816, blue: 0.816 }; // #D0D0D0
const COLOR_ORANGE = { red: 1.0,   green: 0.890, blue: 0.710 }; // #FFE3B5

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
  pipeline?: string;
  dealAmount?: number;
  winProbability?: number;
  expectedCloseDate?: string;
  contactName?: string;
  contactEmail?: string;
  contactDepartment?: string;
  lastContactDate?: string;
}

// ==================== ヘッダーマイグレーション ====================

async function migrateHeadersIfNeeded(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetId: number,
  tabTitle: string,
): Promise<void> {
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${tabTitle}'!A1:V1`,
  });
  const currentHeaders = headerRes.data.values?.[0] ?? [];

  // 旧形式1: 7列（ステータスがG列）→ 挿入して8列にする
  if (currentHeaders.length === 7 && currentHeaders[6] === "ステータス") {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          insertDimension: {
            range: { sheetId, dimension: "COLUMNS", startIndex: 6, endIndex: 7 },
            inheritFromBefore: false,
          },
        }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${tabTitle}'!G1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [["フォーム営業文"]] },
    });
    currentHeaders.splice(6, 0, "フォーム営業文");
  }

  // 旧形式2: 8列 → 14列
  if (currentHeaders.length === 8 && currentHeaders[7] === "ステータス") {
    const newHeaders = HEADERS.slice(8, 14);
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${tabTitle}'!I1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [newHeaders] },
    });
    await applyHeaderStyles(sheets, spreadsheetId, sheetId, 8, 14);
    await applyDropdowns(sheets, spreadsheetId, sheetId, 14);
    currentHeaders.push(...newHeaders);
  }

  // 旧形式3: 14列 → 22列
  if (currentHeaders.length === 14 && currentHeaders[13] === "推奨アクション") {
    const newHeaders = HEADERS.slice(14);
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${tabTitle}'!O1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [newHeaders] },
    });
    await applyHeaderStyles(sheets, spreadsheetId, sheetId, 14, COL_COUNT);
    await applyExtendedDropdowns(sheets, spreadsheetId, sheetId);
    await applyConditionalFormatRules(sheets, spreadsheetId, sheetId);
  }
}

/** ヘッダー行のスタイルを指定範囲に適用 */
async function applyHeaderStyles(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetId: number,
  startCol: number,
  endCol: number,
): Promise<void> {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: startCol, endColumnIndex: endCol },
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
        ...COL_WIDTHS.slice(startCol, endCol).map((pixels, i) => ({
          updateDimensionProperties: {
            range: { sheetId, dimension: "COLUMNS", startIndex: startCol + i, endIndex: startCol + i + 1 },
            properties: { pixelSize: pixels },
            fields: "pixelSize",
          },
        })),
      ],
    },
  });
}

/** 基本ドロップダウン（H/J/L列）を設定 */
async function applyDropdowns(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetId: number,
  colCount: number,
): Promise<void> {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // H列: ステータス
        buildDropdownRequest(sheetId, 7, 8, STATUS_OPTIONS),
        // J列: ランク
        buildDropdownRequest(sheetId, 9, 10, RANK_OPTIONS),
        // L列: 接触経路
        buildDropdownRequest(sheetId, 11, 12, CONTACT_PATH_OPTIONS),
      ],
    },
  });
}

/** 拡張列のドロップダウン（O列: パイプライン）を設定 */
async function applyExtendedDropdowns(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetId: number,
): Promise<void> {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // O列: パイプライン
        buildDropdownRequest(sheetId, 14, 15, PIPELINE_STAGES),
        // P列: 通貨書式
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 15, endColumnIndex: 16 },
            cell: {
              userEnteredFormat: {
                numberFormat: { type: "NUMBER", pattern: "#,##0" },
              },
            },
            fields: "userEnteredFormat.numberFormat",
          },
        },
      ],
    },
  });
}

/** 条件付き書式ルールを設定 */
async function applyConditionalFormatRules(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetId: number,
): Promise<void> {
  const rules: object[] = [
    // J列（ランク）の色分け
    buildCondFmtRule(sheetId, 9, 10, "A", COLOR_GREEN),
    buildCondFmtRule(sheetId, 9, 10, "B", COLOR_YELLOW),
    buildCondFmtRule(sheetId, 9, 10, "C", COLOR_RED),
    // H列（ステータス）の色分け
    buildCondFmtRule(sheetId, 7, 8, "Aランク対応中", COLOR_GREEN),
    buildCondFmtRule(sheetId, 7, 8, "未アプローチ", COLOR_RED),
    buildCondFmtRule(sheetId, 7, 8, "ナーチャリング中", COLOR_YELLOW),
    buildCondFmtRule(sheetId, 7, 8, "アプローチ済み", COLOR_ORANGE),
    buildCondFmtRule(sheetId, 7, 8, "フォーム営業完了", COLOR_BLUE),
    // O列（パイプライン）の色分け
    buildCondFmtRule(sheetId, 14, 15, "受注", COLOR_GREEN),
    buildCondFmtRule(sheetId, 14, 15, "失注", COLOR_GREY),
    buildCondFmtRule(sheetId, 14, 15, "商談", COLOR_BLUE),
    buildCondFmtRule(sheetId, 14, 15, "提案", COLOR_BLUE),
    buildCondFmtRule(sheetId, 14, 15, "交渉", COLOR_ORANGE),
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: rules },
  });
}

// ==================== タブ管理 ====================

/** 全タブが存在することを確認し、なければ作成する */
async function ensureAllTabs(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
): Promise<Map<string, number>> {
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
  const existingTabs = new Map<string, number>();
  for (const s of meta.data.sheets ?? []) {
    existingTabs.set(s.properties?.title ?? "", s.properties?.sheetId ?? 0);
  }

  const tabConfigs: Array<{
    name: string;
    headers: string[];
    colWidths: number[];
    dropdowns?: Array<{ startCol: number; endCol: number; options: string[] }>;
    conditionalFormats?: boolean;
  }> = [
    {
      name: TAB_CONTACTS,
      headers: CONTACT_HEADERS,
      colWidths: CONTACT_COL_WIDTHS,
      dropdowns: [{ startCol: 6, endCol: 7, options: KEYMAN_OPTIONS }],
    },
    {
      name: TAB_ACTIVITIES,
      headers: ACTIVITY_HEADERS,
      colWidths: ACTIVITY_COL_WIDTHS,
      dropdowns: [{ startCol: 2, endCol: 3, options: ACTIVITY_TYPES }],
    },
    {
      name: TAB_TASKS,
      headers: TASK_HEADERS,
      colWidths: TASK_COL_WIDTHS,
      dropdowns: [
        { startCol: 4, endCol: 5, options: TASK_PRIORITY },
        { startCol: 5, endCol: 6, options: TASK_STATUS_OPTIONS },
      ],
      conditionalFormats: true,
    },
    {
      name: TAB_DASHBOARD,
      headers: [],
      colWidths: [],
    },
  ];

  for (const cfg of tabConfigs) {
    if (existingTabs.has(cfg.name)) continue;

    // タブを追加
    const addRes = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          addSheet: { properties: { title: cfg.name } },
        }],
      },
    });
    const newSheetId = addRes.data.replies?.[0].addSheet?.properties?.sheetId ?? 0;
    existingTabs.set(cfg.name, newSheetId);

    if (cfg.headers.length === 0) continue; // ダッシュボードはヘッダーなし

    // ヘッダー書き込み
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${cfg.name}'!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [cfg.headers] },
    });

    // スタイル + ドロップダウン
    const requests: object[] = [
      // ヘッダースタイル
      {
        repeatCell: {
          range: { sheetId: newSheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: cfg.headers.length },
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
      // 列幅
      ...cfg.colWidths.map((pixels, i) => ({
        updateDimensionProperties: {
          range: { sheetId: newSheetId, dimension: "COLUMNS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: pixels },
          fields: "pixelSize",
        },
      })),
      // ヘッダー行高さ
      {
        updateDimensionProperties: {
          range: { sheetId: newSheetId, dimension: "ROWS", startIndex: 0, endIndex: 1 },
          properties: { pixelSize: 40 },
          fields: "pixelSize",
        },
      },
      // ヘッダー行固定
      {
        updateSheetProperties: {
          properties: { sheetId: newSheetId, gridProperties: { frozenRowCount: 1 } },
          fields: "gridProperties.frozenRowCount",
        },
      },
    ];

    // ドロップダウン
    for (const dd of cfg.dropdowns ?? []) {
      requests.push(buildDropdownRequest(newSheetId, dd.startCol, dd.endCol, dd.options));
    }

    // タスクタブ用条件付き書式
    if (cfg.conditionalFormats) {
      // 優先度「高」= 赤背景
      requests.push(buildCondFmtRule(newSheetId, 4, 5, "高", COLOR_RED));
      // ステータス「完了」= グレー（行全体）
      requests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{ sheetId: newSheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 0, endColumnIndex: cfg.headers.length }],
            booleanRule: {
              condition: { type: "CUSTOM_FORMULA", values: [{ userEnteredValue: `=$F2="完了"` }] },
              format: { backgroundColor: COLOR_GREY, textFormat: { foregroundColor: { red: 0.5, green: 0.5, blue: 0.5 } } },
            },
          },
          index: 0,
        },
      });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });
  }

  return existingTabs;
}

// ==================== ユーティリティ ====================

function buildDropdownRequest(sheetId: number, startCol: number, endCol: number, options: string[]): object {
  return {
    setDataValidation: {
      range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: startCol, endColumnIndex: endCol },
      rule: {
        condition: {
          type: "ONE_OF_LIST",
          values: options.map((v) => ({ userEnteredValue: v })),
        },
        showCustomUi: true,
        strict: false,
      },
    },
  };
}

function buildCondFmtRule(sheetId: number, startCol: number, endCol: number, value: string, bgColor: object): object {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [{ sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: startCol, endColumnIndex: endCol }],
        booleanRule: {
          condition: { type: "TEXT_EQ", values: [{ userEnteredValue: value }] },
          format: { backgroundColor: bgColor },
        },
      },
      index: 0,
    },
  };
}

// ==================== メイン関数 ====================

/** フォルダ内で管理シートを検索し、なければ新規作成して { spreadsheetId, sheetId, tabTitle } を返す */
async function findOrCreateTrackingSheet(
  drive: drive_v3.Drive,
  sheets: sheets_v4.Sheets,
  folderId: string,
): Promise<{ spreadsheetId: string; sheetId: number; tabTitle: string }> {
  const listRes = await drive.files.list({
    q: `name='${TRACKING_SHEET_NAME}' and '${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: "files(id)",
    pageSize: 1,
  });

  const existing = listRes.data.files?.[0];
  if (existing?.id) {
    const spreadsheetId = existing.id;
    const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
    const firstSheet = meta.data.sheets?.[0].properties;
    const sheetId = firstSheet?.sheetId ?? 0;
    const tabTitle = firstSheet?.title ?? SHEET_TAB;

    // ヘッダーの自動マイグレーション
    await migrateHeadersIfNeeded(sheets, spreadsheetId, sheetId, tabTitle);

    // 追加タブの自動作成
    await ensureAllTabs(sheets, spreadsheetId);

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

  // ヘッダーのスタイル + ドロップダウン + 条件付き書式
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
        // ドロップダウン
        buildDropdownRequest(sheetId, 7, 8, STATUS_OPTIONS),
        buildDropdownRequest(sheetId, 9, 10, RANK_OPTIONS),
        buildDropdownRequest(sheetId, 11, 12, CONTACT_PATH_OPTIONS),
        buildDropdownRequest(sheetId, 14, 15, PIPELINE_STAGES),
        // P列: 通貨書式
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 1, endRowIndex: 10000, startColumnIndex: 15, endColumnIndex: 16 },
            cell: {
              userEnteredFormat: {
                numberFormat: { type: "NUMBER", pattern: "#,##0" },
              },
            },
            fields: "userEnteredFormat.numberFormat",
          },
        },
      ],
    },
  });

  // 条件付き書式
  await applyConditionalFormatRules(sheets, spreadsheetId, sheetId);

  // 追加タブの自動作成
  await ensureAllTabs(sheets, spreadsheetId);

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
    range: `'${tabTitle}'!A:V`,
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
        row.pipeline ?? "リード",
        row.dealAmount ?? "",
        row.winProbability ?? "",
        row.expectedCloseDate ?? "",
        row.contactName ?? "",
        row.contactEmail ?? "",
        row.contactDepartment ?? "",
        row.lastContactDate ?? "",
      ]],
    },
  });

  // 追加された行のインデックスを取得
  const updatedRange = appendRes.data.updates?.updatedRange ?? "";
  const match = updatedRange.match(/(\d+)(?::.*)?$/);
  const rowIndex = match ? parseInt(match[1], 10) - 1 : -1;
  if (rowIndex < 1) return;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // 行全体の背景色を白・文字色を黒にリセット
        {
          repeatCell: {
            range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
            cell: {
              userEnteredFormat: {
                backgroundColor: WHITE,
                textFormat: { foregroundColor: BLACK, bold: false },
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

  // アクティビティを自動記録
  try {
    await appendActivityRow(sheets, spreadsheetId, {
      companyName: row.companyName,
      activityType: "フォーム営業",
      content: `レポート生成・CRM登録: ${row.reportUrl}`,
      recorder: "自動",
    });
  } catch (e) {
    // アクティビティ記録失敗はメイン処理をブロックしない
    console.error("アクティビティ記録に失敗:", e);
  }
}

// ==================== アクティビティ記録 ====================

/** アクティビティタブに1行追加する */
export async function appendActivityRow(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  args: {
    companyName: string;
    activityType: string;
    contactName?: string;
    content: string;
    result?: string;
    recorder?: string;
  },
): Promise<void> {
  // アクティビティタブの存在を確認
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
  let activityTab: string | null = null;
  for (const s of meta.data.sheets ?? []) {
    if (s.properties?.title === TAB_ACTIVITIES) {
      activityTab = TAB_ACTIVITIES;
      break;
    }
  }
  if (!activityTab) return; // タブが存在しない場合はスキップ

  const now = new Date();
  const timestamp = `${String(now.getFullYear()).slice(2)}_${String(now.getMonth() + 1).padStart(2, "0")}_${String(now.getDate()).padStart(2, "0")} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `'${TAB_ACTIVITIES}'!A:G`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[
        timestamp,
        args.companyName,
        args.activityType,
        args.contactName ?? "",
        args.content,
        args.result ?? "",
        args.recorder ?? "手動",
      ]],
    },
  });

  // リストタブのV列（最終接触日）も更新
  try {
    const listTab = meta.data.sheets?.[0].properties?.title ?? SHEET_TAB;
    const dataRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${listTab}'!B:B`,
    });
    const rows = dataRes.data.values ?? [];
    for (let i = rows.length - 1; i >= 1; i--) {
      const cell = String(rows[i][0] ?? "");
      if (cell.includes(args.companyName) || args.companyName.includes(cell)) {
        const dateStr = `${String(now.getFullYear()).slice(2)}_${String(now.getMonth() + 1).padStart(2, "0")}_${String(now.getDate()).padStart(2, "0")}`;
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `'${listTab}'!V${i + 1}`,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: [[dateStr]] },
        });
        break;
      }
    }
  } catch (_) {
    // 最終接触日の更新失敗は無視
  }
}

// ==================== エクスポート ====================

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
    pipeline: "リード",
  });
}
