/**
 * Google スプレッドシートに 4タブ＋前提条件・承認 を一括書き込み
 * 環境変数 GOOGLE_APPLICATION_CREDENTIALS にサービスアカウントJSONのパスを指定する。
 */

import fs from "node:fs";
import path from "node:path";
import { google } from "googleapis";
import type { ProposalSheet } from "./types.js";
import {
  costEffectToSheetRows,
  costEffectToVisualRows,
  packageToSheetRows,
  packageToVisualRows,
  roadmapToSheetRows,
  roadmapToVisualRows,
  premiseToSheetRows,
} from "./build-four-tabs.js";
import type { CostEffectVisualRow, PackageVisualRow, RoadmapVisualRow } from "./build-four-tabs.js";
import type { CostEffectRow, PackageRow, RoadmapRow, PremiseRow } from "./types.js";

const SHEET_TITLES = [
  "考えられる施策",
  "費用対効果など",
  "パッケージ",
  "ロードマップ",
  "前提条件・承認",
] as const;

function proposalToSheetRows(sheet: ProposalSheet): string[][] {
  const header = ["ブロック種別", "施策名", "項目", "値", "単位", "メモ"];
  const rows: string[][] = [header];
  for (const r of sheet.common.rows) {
    rows.push(["共通", "", r.item, r.value, r.unit, r.memo]);
  }
  for (const inv of sheet.initiatives) {
    for (const r of inv.rows) {
      rows.push(["施策", inv.name, r.item, r.value, r.unit, r.memo]);
    }
  }
  return rows;
}

type RowType = "title" | "empty" | "section" | "tableHeader" | "data";
interface VisualRow { row: string[]; type: RowType; }

function proposalToVisualRows(sheet: ProposalSheet): VisualRow[] {
  const result: VisualRow[] = [];
  result.push({ row: ["考えられる施策", "", "", ""], type: "title" });
  result.push({ row: ["", "", "", ""], type: "empty" });
  result.push({ row: ["共通", "", "", ""], type: "section" });
  result.push({ row: ["項目", "値", "単位", "メモ"], type: "tableHeader" });
  for (const r of sheet.common.rows) {
    result.push({ row: [r.item, r.value, r.unit, r.memo], type: "data" });
  }
  for (const inv of sheet.initiatives) {
    result.push({ row: ["", "", "", ""], type: "empty" });
    result.push({ row: [inv.name, "", "", ""], type: "section" });
    result.push({ row: ["項目", "値", "単位", "メモ"], type: "tableHeader" });
    for (const r of inv.rows) {
      result.push({ row: [r.item, r.value, r.unit, r.memo], type: "data" });
    }
  }
  return result;
}

async function applyProposalFormatting(
  sheetsClient: ReturnType<typeof google.sheets>,
  spreadsheetId: string,
  sheetId: number,
  visualRows: VisualRow[],
) {
  const requests: object[] = [];

  for (const [idx, size] of [[0, 250], [1, 130], [2, 80], [3, 350]] as [number, number][]) {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId, dimension: "COLUMNS", startIndex: idx, endIndex: idx + 1 },
        properties: { pixelSize: size },
        fields: "pixelSize",
      },
    });
  }

  for (let i = 0; i < visualRows.length; i++) {
    const { type } = visualRows[i];
    if (type === "title") {
      requests.push({
        mergeCells: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 4 },
          mergeType: "MERGE_ALL",
        },
      });
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 4 },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.118, green: 0.227, blue: 0.376 },
              textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 13, bold: true },
              verticalAlignment: "MIDDLE",
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)",
        },
      });
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 40 },
          fields: "pixelSize",
        },
      });
    } else if (type === "section") {
      requests.push({
        mergeCells: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 4 },
          mergeType: "MERGE_ALL",
        },
      });
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 4 },
          cell: {
            userEnteredFormat: {
              textFormat: { foregroundColor: { red: 0.102, green: 0.337, blue: 0.675 }, fontSize: 11, bold: true },
            },
          },
          fields: "userEnteredFormat(textFormat)",
        },
      });
    } else if (type === "tableHeader") {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 4 },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.788, green: 0.855, blue: 0.937 },
              textFormat: { bold: true, fontSize: 10 },
              horizontalAlignment: "CENTER",
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
        },
      });
    } else if (type === "data") {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 1, endColumnIndex: 2 },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 1.0, green: 0.973, blue: 0.878 },
              horizontalAlignment: "RIGHT",
            },
          },
          fields: "userEnteredFormat(backgroundColor,horizontalAlignment)",
        },
      });
    }
  }

  if (requests.length > 0) {
    await sheetsClient.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  }
}

async function applyCostEffectFormatting(
  sheetsClient: ReturnType<typeof google.sheets>,
  spreadsheetId: string,
  sheetId: number,
  visualRows: CostEffectVisualRow[],
) {
  const requests: object[] = [];

  // 列幅: A=250, B=150, C=110, D=110, E=90, F=100, G=220, H=220, I=150
  for (const [idx, size] of [[0, 250], [1, 150], [2, 110], [3, 110], [4, 90], [5, 100], [6, 220], [7, 220], [8, 150]] as [number, number][]) {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId, dimension: "COLUMNS", startIndex: idx, endIndex: idx + 1 },
        properties: { pixelSize: size },
        fields: "pixelSize",
      },
    });
  }

  const NAVY = { red: 0.118, green: 0.227, blue: 0.376 };
  const WHITE = { red: 1, green: 1, blue: 1 };
  const LIGHT_BLUE = { red: 0.851, green: 0.918, blue: 0.988 };
  const COL_COUNT = 9;

  for (let i = 0; i < visualRows.length; i++) {
    const { type } = visualRows[i];

    if (type === "title") {
      requests.push({
        mergeCells: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          mergeType: "MERGE_ALL",
        },
      });
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          cell: {
            userEnteredFormat: {
              backgroundColor: NAVY,
              textFormat: { foregroundColor: WHITE, fontSize: 14, bold: true },
              verticalAlignment: "MIDDLE",
              padding: { left: 12 },
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)",
        },
      });
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 45 },
          fields: "pixelSize",
        },
      });
    } else if (type === "tableHeader") {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          cell: {
            userEnteredFormat: {
              backgroundColor: NAVY,
              textFormat: { foregroundColor: WHITE, fontSize: 10, bold: true },
              horizontalAlignment: "CENTER",
              verticalAlignment: "MIDDLE",
              wrapStrategy: "WRAP",
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)",
        },
      });
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 50 },
          fields: "pixelSize",
        },
      });
    } else if (type === "data") {
      // 行の高さ
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 45 },
          fields: "pixelSize",
        },
      });
      // A列: テキスト折り返し
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 1 },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(wrapStrategy,verticalAlignment)",
        },
      });
      // C列(月次効果): ¥ 書式 + 右揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 2, endColumnIndex: 3 },
          cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: "¥#,##0" }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
      // D列(年次効果): ¥ 書式 + 右揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 3, endColumnIndex: 4 },
          cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: "¥#,##0" }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
      // E列(回収期間): 数値書式 + 右揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 4, endColumnIndex: 5 },
          cell: { userEnteredFormat: { numberFormat: { type: "NUMBER", pattern: "#,##0.0" }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
      // F列(年次ROI): % 書式 + 右揃え (値は89.0のような数値)
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 5, endColumnIndex: 6 },
          cell: { userEnteredFormat: { numberFormat: { type: "NUMBER", pattern: '0.0"%"' }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
    } else if (type === "total") {
      // A-B をマージ
      requests.push({
        mergeCells: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 2 },
          mergeType: "MERGE_ALL",
        },
      });
      // 行全体: 薄青背景・太字
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          cell: { userEnteredFormat: { backgroundColor: LIGHT_BLUE, textFormat: { bold: true } } },
          fields: "userEnteredFormat(backgroundColor,textFormat)",
        },
      });
      // C列: ¥ 書式
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 2, endColumnIndex: 3 },
          cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: "¥#,##0" }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
      // D列: ¥ 書式
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 3, endColumnIndex: 4 },
          cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: "¥#,##0" }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
      // E列: 数値書式
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 4, endColumnIndex: 5 },
          cell: { userEnteredFormat: { numberFormat: { type: "NUMBER", pattern: "#,##0.0" }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
      // F列: % 書式
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 5, endColumnIndex: 6 },
          cell: { userEnteredFormat: { numberFormat: { type: "NUMBER", pattern: '0.0"%"' }, horizontalAlignment: "RIGHT" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment)",
        },
      });
    }
  }

  if (requests.length > 0) {
    await sheetsClient.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  }
}

async function applyPackageFormatting(
  sheetsClient: ReturnType<typeof google.sheets>,
  spreadsheetId: string,
  sheetId: number,
  visualRows: PackageVisualRow[],
) {
  const requests: object[] = [];

  // 列幅: A=150, B=250, C=200, D=80, E=130, F=120, G=100, H=200
  for (const [idx, size] of [[0, 150], [1, 250], [2, 200], [3, 80], [4, 130], [5, 120], [6, 100], [7, 200]] as [number, number][]) {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId, dimension: "COLUMNS", startIndex: idx, endIndex: idx + 1 },
        properties: { pixelSize: size },
        fields: "pixelSize",
      },
    });
  }

  const NAVY = { red: 0.118, green: 0.227, blue: 0.376 };
  const WHITE = { red: 1, green: 1, blue: 1 };
  const LIGHT_BLUE_BG = { red: 0.788, green: 0.855, blue: 0.937 };
  const COL_COUNT = 8;

  for (let i = 0; i < visualRows.length; i++) {
    const { type } = visualRows[i];

    if (type === "title") {
      requests.push({
        mergeCells: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          mergeType: "MERGE_ALL",
        },
      });
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          cell: {
            userEnteredFormat: {
              backgroundColor: NAVY,
              textFormat: { foregroundColor: WHITE, fontSize: 14, bold: true },
              verticalAlignment: "MIDDLE",
              padding: { left: 12 },
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)",
        },
      });
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 45 },
          fields: "pixelSize",
        },
      });
    } else if (type === "empty") {
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 20 },
          fields: "pixelSize",
        },
      });
    } else if (type === "tableHeader") {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          cell: {
            userEnteredFormat: {
              backgroundColor: LIGHT_BLUE_BG,
              textFormat: { bold: true, fontSize: 10 },
              horizontalAlignment: "CENTER",
              verticalAlignment: "MIDDLE",
              wrapStrategy: "WRAP",
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)",
        },
      });
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 45 },
          fields: "pixelSize",
        },
      });
    } else if (type === "data") {
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 60 },
          fields: "pixelSize",
        },
      });
      // A/B/C列: テキスト折り返し
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 3 },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(wrapStrategy,verticalAlignment)",
        },
      });
      // D列(システム数): 中央揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 3, endColumnIndex: 4 },
          cell: { userEnteredFormat: { horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(horizontalAlignment,verticalAlignment)",
        },
      });
      // E列(初期費用): ¥ 書式 + 右揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 4, endColumnIndex: 5 },
          cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: "¥#,##0" }, horizontalAlignment: "RIGHT", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment,verticalAlignment)",
        },
      });
      // F列(月次効果): ¥ 書式 + 右揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 5, endColumnIndex: 6 },
          cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: "¥#,##0" }, horizontalAlignment: "RIGHT", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment,verticalAlignment)",
        },
      });
      // G列(回収期間): 小数書式 + 右揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 6, endColumnIndex: 7 },
          cell: { userEnteredFormat: { numberFormat: { type: "NUMBER", pattern: "#,##0.0" }, horizontalAlignment: "RIGHT", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(numberFormat,horizontalAlignment,verticalAlignment)",
        },
      });
      // H列(メモ): 折り返し
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 7, endColumnIndex: 8 },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(wrapStrategy,verticalAlignment)",
        },
      });
    }
  }

  if (requests.length > 0) {
    await sheetsClient.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  }
}

async function applyRoadmapFormatting(
  sheetsClient: ReturnType<typeof google.sheets>,
  spreadsheetId: string,
  sheetId: number,
  visualRows: RoadmapVisualRow[],
) {
  const requests: object[] = [];

  // 列幅: A=130, B=350, C-J=55px each
  requests.push({
    updateDimensionProperties: {
      range: { sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 130 },
      fields: "pixelSize",
    },
  });
  requests.push({
    updateDimensionProperties: {
      range: { sheetId, dimension: "COLUMNS", startIndex: 1, endIndex: 2 },
      properties: { pixelSize: 350 },
      fields: "pixelSize",
    },
  });
  requests.push({
    updateDimensionProperties: {
      range: { sheetId, dimension: "COLUMNS", startIndex: 2, endIndex: 10 },
      properties: { pixelSize: 55 },
      fields: "pixelSize",
    },
  });

  const NAVY = { red: 0.118, green: 0.227, blue: 0.376 };
  const WHITE = { red: 1, green: 1, blue: 1 };
  const LIGHT_BLUE_BG = { red: 0.788, green: 0.855, blue: 0.937 };
  const WEEK_FILL = { red: 0.647, green: 0.749, blue: 0.875 };
  const COL_COUNT = 10;

  for (let i = 0; i < visualRows.length; i++) {
    const { type, weekMask } = visualRows[i];

    if (type === "title") {
      requests.push({
        mergeCells: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          mergeType: "MERGE_ALL",
        },
      });
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          cell: {
            userEnteredFormat: {
              backgroundColor: NAVY,
              textFormat: { foregroundColor: WHITE, fontSize: 14, bold: true },
              verticalAlignment: "MIDDLE",
              padding: { left: 12 },
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)",
        },
      });
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 45 },
          fields: "pixelSize",
        },
      });
    } else if (type === "empty") {
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 20 },
          fields: "pixelSize",
        },
      });
    } else if (type === "tableHeader") {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: COL_COUNT },
          cell: {
            userEnteredFormat: {
              backgroundColor: LIGHT_BLUE_BG,
              textFormat: { bold: true, fontSize: 10 },
              horizontalAlignment: "CENTER",
              verticalAlignment: "MIDDLE",
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
        },
      });
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 40 },
          fields: "pixelSize",
        },
      });
    } else if (type === "data") {
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "ROWS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize: 50 },
          fields: "pixelSize",
        },
      });
      // A列: 折り返し・中央揃え
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 0, endColumnIndex: 1 },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(wrapStrategy,verticalAlignment)",
        },
      });
      // B列: 折り返し
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: 1, endColumnIndex: 2 },
          cell: { userEnteredFormat: { wrapStrategy: "WRAP", verticalAlignment: "MIDDLE" } },
          fields: "userEnteredFormat(wrapStrategy,verticalAlignment)",
        },
      });
      // 週列(C-J): weekMask に従って塗りつぶし
      if (weekMask) {
        for (let w = 0; w < 8; w++) {
          const colIndex = 2 + w;
          const bg = weekMask[w] ? WEEK_FILL : { red: 1, green: 1, blue: 1 };
          requests.push({
            repeatCell: {
              range: { sheetId, startRowIndex: i, endRowIndex: i + 1, startColumnIndex: colIndex, endColumnIndex: colIndex + 1 },
              cell: { userEnteredFormat: { backgroundColor: bg } },
              fields: "userEnteredFormat(backgroundColor)",
            },
          });
        }
      }
    }
  }

  if (requests.length > 0) {
    await sheetsClient.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  }
}

export interface FullSpreadsheetData {
  proposal: ProposalSheet;
  costEffect: CostEffectRow[];
  packages: PackageRow[];
  roadmap: RoadmapRow[];
  premise: PremiseRow[];
}

/**
 * 1つのスプレッドシートを新規作成し、5シートにデータを書き込む。
 * @returns 作成したスプレッドシートのURL
 */
export async function writeToGoogleSheets(
  title: string,
  data: FullSpreadsheetData,
  credentialsPath?: string,
): Promise<string> {
  const keyPath = credentialsPath || process.env.GOOGLE_APPLICATION_CREDENTIALS;
  if (!keyPath || !fs.existsSync(keyPath)) {
    throw new Error(
      "GOOGLE_APPLICATION_CREDENTIALS にサービスアカウントJSONのパスを設定してください。",
    );
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

  const sheets = google.sheets({ version: "v4", auth });

  const createRes = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title },
      sheets: SHEET_TITLES.map((t) => ({ properties: { title: t } })),
    },
  });

  const spreadsheetId = createRes.data.spreadsheetId;
  if (!spreadsheetId) throw new Error("Failed to create spreadsheet");

  const sheetIds = createRes.data.sheets?.map((s) => s.properties?.sheetId ?? 0) ?? [];

  const proposalVisual = proposalToVisualRows(data.proposal);
  const costEffectVisual = costEffectToVisualRows(data.costEffect);
  const packageVisual = packageToVisualRows(data.packages);
  const roadmapVisual = roadmapToVisualRows(data.roadmap);
  const premiseRows = premiseToSheetRows(data.premise);

  const allData = [
    proposalVisual.map((r) => r.row),
    costEffectVisual.map((r) => r.row),
    packageVisual.map((r) => r.row),
    roadmapVisual.map((r) => r.row),
    premiseRows,
  ];

  const updates = SHEET_TITLES.map((sheetTitle, i) => ({
    range: `'${sheetTitle}'!A1`,
    values: allData[i],
  }));

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: updates.map((u) => ({ range: u.range, values: u.values })),
    },
  });

  if (sheetIds.length > 0) {
    await applyProposalFormatting(sheets, spreadsheetId, sheetIds[0], proposalVisual);
  }
  if (sheetIds.length > 1) {
    await applyCostEffectFormatting(sheets, spreadsheetId, sheetIds[1], costEffectVisual);
  }
  if (sheetIds.length > 2) {
    await applyPackageFormatting(sheets, spreadsheetId, sheetIds[2], packageVisual);
  }
  if (sheetIds.length > 3) {
    await applyRoadmapFormatting(sheets, spreadsheetId, sheetIds[3], roadmapVisual);
  }

  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
}

/** 認証がない場合のフォールバック：5つのCSVを data/out に出力 */
export function writeToCsvFallback(
  data: FullSpreadsheetData,
  outDir: string,
  prefix: string,
  writeCsv: (headers: string[], rows: string[][], path: string) => void,
): string[] {
  const paths: string[] = [];
  const proposalRows = proposalToSheetRows(data.proposal);
  const proposalPath = path.join(outDir, `${prefix}_考えられる施策.csv`);
  writeCsv(proposalRows[0], proposalRows.slice(1), proposalPath);
  paths.push(proposalPath);

  const costEffectRows = costEffectToSheetRows(data.costEffect);
  const costPath = path.join(outDir, `${prefix}_費用対効果など.csv`);
  writeCsv(costEffectRows[0], costEffectRows.slice(1), costPath);
  paths.push(costPath);

  const packageRows = packageToSheetRows(data.packages);
  const packagePath = path.join(outDir, `${prefix}_パッケージ.csv`);
  writeCsv(packageRows[0], packageRows.slice(1), packagePath);
  paths.push(packagePath);

  const roadmapRows = roadmapToSheetRows(data.roadmap);
  const roadmapPath = path.join(outDir, `${prefix}_ロードマップ.csv`);
  writeCsv(roadmapRows[0], roadmapRows.slice(1), roadmapPath);
  paths.push(roadmapPath);

  const premiseRows = premiseToSheetRows(data.premise);
  const premisePath = path.join(outDir, `${prefix}_前提条件・承認.csv`);
  writeCsv(premiseRows[0], premiseRows.slice(1), premisePath);
  paths.push(premisePath);

  return paths;
}
