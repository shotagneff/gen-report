#!/usr/bin/env npx tsx
/**
 * ダッシュボードタブの初期化スクリプト。
 * KPI数式の書き込みとグラフの作成を行う。
 * Usage:
 *   npx tsx scripts/init-crm-dashboard.ts           # ダッシュボードを初期化
 *   npx tsx scripts/init-crm-dashboard.ts --refresh  # グラフを再作成
 */

import "dotenv/config";
import { getCRMConnection } from "../src/crm-common.js";

const TAB_DASHBOARD = "ダッシュボード";

function parseArgs(argv: string[]): { refresh: boolean } {
  let refresh = false;
  for (const arg of argv) {
    if (arg === "--refresh") refresh = true;
    else if (arg === "-h" || arg === "--help") {
      console.log(`
Usage: npx tsx scripts/init-crm-dashboard.ts [--refresh]

  リード管理CRMのダッシュボードタブにKPI数式とグラフを設定します。

  --refresh   既存のグラフを削除して再作成
`);
      process.exit(0);
    }
  }
  return { refresh };
}

async function main(): Promise<void> {
  const { refresh } = parseArgs(process.argv.slice(2));

  const conn = await getCRMConnection();
  const dashTab = conn.tabs.get(TAB_DASHBOARD);

  if (!dashTab) {
    console.error("Error: ダッシュボードタブが見つかりません。gen-reportを一度実行してCRMタブを作成してください。");
    process.exit(1);
  }

  const sheetId = dashTab!.sheetId;

  // 既存チャートを削除（--refresh時）
  if (refresh) {
    const meta = await conn.sheets.spreadsheets.get({
      spreadsheetId: conn.spreadsheetId,
      fields: "sheets(properties,charts)",
    });
    const dashSheet = meta.data.sheets?.find(s => s.properties?.sheetId === sheetId);
    const charts = dashSheet?.charts ?? [];
    if (charts.length > 0) {
      await conn.sheets.spreadsheets.batchUpdate({
        spreadsheetId: conn.spreadsheetId,
        requestBody: {
          requests: charts.map(c => ({
            deleteEmbeddedObject: { objectId: c.chartId },
          })),
        },
      });
      console.error(`既存グラフ ${charts.length} 件を削除しました`);
    }
  }

  // KPI数式を書き込み
  const kpiData = [
    // セクション1: 概要KPI（行1〜5）
    ["リード管理ダッシュボード", "", "", "", ""],
    ["", "", "", "", ""],
    ["総リード数", "Aランク", "Bランク", "Cランク", "未スコア"],
    [
      "=COUNTA('リスト'!B:B)-1",
      "=COUNTIF('リスト'!J:J,\"A\")",
      "=COUNTIF('リスト'!J:J,\"B\")",
      "=COUNTIF('リスト'!J:J,\"C\")",
      "=COUNTIF('リスト'!J:J,\"\")-1",
    ],
    ["", "", "", "", ""],
    // セクション2: パイプライン集計（行6〜14）
    ["パイプライン別集計", "", "", ""],
    ["ステージ", "件数", "金額合計", "加重金額"],
    ["リード", "=COUNTIF('リスト'!O:O,\"リード\")", "=SUMIF('リスト'!O:O,\"リード\",'リスト'!P:P)", "=SUMPRODUCT(('リスト'!O2:O10000=\"リード\")*('リスト'!P2:P10000)*('リスト'!Q2:Q10000/100))"],
    ["アプローチ中", "=COUNTIF('リスト'!O:O,\"アプローチ中\")", "=SUMIF('リスト'!O:O,\"アプローチ中\",'リスト'!P:P)", "=SUMPRODUCT(('リスト'!O2:O10000=\"アプローチ中\")*('リスト'!P2:P10000)*('リスト'!Q2:Q10000/100))"],
    ["商談", "=COUNTIF('リスト'!O:O,\"商談\")", "=SUMIF('リスト'!O:O,\"商談\",'リスト'!P:P)", "=SUMPRODUCT(('リスト'!O2:O10000=\"商談\")*('リスト'!P2:P10000)*('リスト'!Q2:Q10000/100))"],
    ["提案", "=COUNTIF('リスト'!O:O,\"提案\")", "=SUMIF('リスト'!O:O,\"提案\",'リスト'!P:P)", "=SUMPRODUCT(('リスト'!O2:O10000=\"提案\")*('リスト'!P2:P10000)*('リスト'!Q2:Q10000/100))"],
    ["交渉", "=COUNTIF('リスト'!O:O,\"交渉\")", "=SUMIF('リスト'!O:O,\"交渉\",'リスト'!P:P)", "=SUMPRODUCT(('リスト'!O2:O10000=\"交渉\")*('リスト'!P2:P10000)*('リスト'!Q2:Q10000/100))"],
    ["受注", "=COUNTIF('リスト'!O:O,\"受注\")", "=SUMIF('リスト'!O:O,\"受注\",'リスト'!P:P)", ""],
    ["失注", "=COUNTIF('リスト'!O:O,\"失注\")", "=SUMIF('リスト'!O:O,\"失注\",'リスト'!P:P)", ""],
    ["", "", "", ""],
    // セクション3: ステータス集計（行16〜23）
    ["ステータス別集計", "", ""],
    ["ステータス", "件数", ""],
    ["未アプローチ", "=COUNTIF('リスト'!H:H,\"未アプローチ\")", ""],
    ["アプローチ済み", "=COUNTIF('リスト'!H:H,\"アプローチ済み\")", ""],
    ["フォーム営業完了", "=COUNTIF('リスト'!H:H,\"フォーム営業完了\")", ""],
    ["Aランク対応中", "=COUNTIF('リスト'!H:H,\"Aランク対応中\")", ""],
    ["ナーチャリング中", "=COUNTIF('リスト'!H:H,\"ナーチャリング中\")", ""],
    ["3ヶ月後フォロー", "=COUNTIF('リスト'!H:H,\"3ヶ月後フォロー\")", ""],
  ];

  await conn.sheets.spreadsheets.values.update({
    spreadsheetId: conn.spreadsheetId,
    range: `'${TAB_DASHBOARD}'!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: kpiData },
  });

  // スタイル設定
  const HEADER_BG = { red: 0.118, green: 0.227, blue: 0.376 };
  const WHITE = { red: 1, green: 1, blue: 1 };
  const LIGHT_BG = { red: 0.937, green: 0.953, blue: 0.976 };

  await conn.sheets.spreadsheets.batchUpdate({
    spreadsheetId: conn.spreadsheetId,
    requestBody: {
      requests: [
        // タイトル行
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 5 },
            cell: {
              userEnteredFormat: {
                backgroundColor: HEADER_BG,
                textFormat: { foregroundColor: WHITE, bold: true, fontSize: 14 },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        },
        // KPIラベル行（行3）
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 5 },
            cell: {
              userEnteredFormat: {
                backgroundColor: LIGHT_BG,
                textFormat: { bold: true },
                horizontalAlignment: "CENTER",
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
          },
        },
        // KPI値行（行4）
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 5 },
            cell: {
              userEnteredFormat: {
                textFormat: { fontSize: 20, bold: true },
                horizontalAlignment: "CENTER",
              },
            },
            fields: "userEnteredFormat(textFormat,horizontalAlignment)",
          },
        },
        // セクションヘッダー（行6, 16）
        ...[5, 16].map(row => ({
          repeatCell: {
            range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: 0, endColumnIndex: 4 },
            cell: {
              userEnteredFormat: {
                backgroundColor: HEADER_BG,
                textFormat: { foregroundColor: WHITE, bold: true },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        })),
        // テーブルヘッダー（行7, 17）
        ...[6, 17].map(row => ({
          repeatCell: {
            range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: 0, endColumnIndex: 4 },
            cell: {
              userEnteredFormat: {
                backgroundColor: LIGHT_BG,
                textFormat: { bold: true },
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat)",
          },
        })),
        // 列幅
        ...[200, 80, 120, 120, 80].map((pixels, i) => ({
          updateDimensionProperties: {
            range: { sheetId, dimension: "COLUMNS", startIndex: i, endIndex: i + 1 },
            properties: { pixelSize: pixels },
            fields: "pixelSize",
          },
        })),
        // 通貨書式（C,D列）
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 7, endRowIndex: 15, startColumnIndex: 2, endColumnIndex: 4 },
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

  // グラフ作成
  await conn.sheets.spreadsheets.batchUpdate({
    spreadsheetId: conn.spreadsheetId,
    requestBody: {
      requests: [
        // グラフ1: パイプラインファネル（横棒グラフ）
        {
          addChart: {
            chart: {
              position: {
                overlayPosition: {
                  anchorCell: { sheetId, rowIndex: 0, columnIndex: 5 },
                  widthPixels: 500,
                  heightPixels: 300,
                },
              },
              spec: {
                title: "パイプライン別件数",
                basicChart: {
                  chartType: "BAR",
                  legendPosition: "NO_LEGEND",
                  axis: [
                    { position: "BOTTOM_AXIS", title: "件数" },
                    { position: "LEFT_AXIS", title: "" },
                  ],
                  domains: [{
                    domain: {
                      sourceRange: { sources: [{ sheetId, startRowIndex: 7, endRowIndex: 14, startColumnIndex: 0, endColumnIndex: 1 }] },
                    },
                  }],
                  series: [{
                    series: {
                      sourceRange: { sources: [{ sheetId, startRowIndex: 7, endRowIndex: 14, startColumnIndex: 1, endColumnIndex: 2 }] },
                    },
                    color: { red: 0.259, green: 0.522, blue: 0.957 },
                  }],
                },
              },
            },
          },
        },
        // グラフ2: ランク分布（ドーナツチャート）
        {
          addChart: {
            chart: {
              position: {
                overlayPosition: {
                  anchorCell: { sheetId, rowIndex: 0, columnIndex: 10 },
                  widthPixels: 400,
                  heightPixels: 300,
                },
              },
              spec: {
                title: "ランク分布",
                pieChart: {
                  legendPosition: "RIGHT_LEGEND",
                  pieHole: 0.4,
                  domain: {
                    sourceRange: { sources: [{ sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 1, endColumnIndex: 4 }] },
                  },
                  series: {
                    sourceRange: { sources: [{ sheetId, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 1, endColumnIndex: 4 }] },
                  },
                },
              },
            },
          },
        },
        // グラフ3: ステータス分布（横棒グラフ）
        {
          addChart: {
            chart: {
              position: {
                overlayPosition: {
                  anchorCell: { sheetId, rowIndex: 16, columnIndex: 5 },
                  widthPixels: 500,
                  heightPixels: 300,
                },
              },
              spec: {
                title: "ステータス別件数",
                basicChart: {
                  chartType: "BAR",
                  legendPosition: "NO_LEGEND",
                  axis: [
                    { position: "BOTTOM_AXIS", title: "件数" },
                    { position: "LEFT_AXIS", title: "" },
                  ],
                  domains: [{
                    domain: {
                      sourceRange: { sources: [{ sheetId, startRowIndex: 18, endRowIndex: 24, startColumnIndex: 0, endColumnIndex: 1 }] },
                    },
                  }],
                  series: [{
                    series: {
                      sourceRange: { sources: [{ sheetId, startRowIndex: 18, endRowIndex: 24, startColumnIndex: 1, endColumnIndex: 2 }] },
                    },
                    color: { red: 0.204, green: 0.659, blue: 0.325 },
                  }],
                },
              },
            },
          },
        },
        // グラフ4: パイプライン金額（積み上げ棒グラフ）
        {
          addChart: {
            chart: {
              position: {
                overlayPosition: {
                  anchorCell: { sheetId, rowIndex: 16, columnIndex: 10 },
                  widthPixels: 500,
                  heightPixels: 300,
                },
              },
              spec: {
                title: "パイプライン別金額",
                basicChart: {
                  chartType: "COLUMN",
                  legendPosition: "BOTTOM_LEGEND",
                  stackedType: "STACKED",
                  axis: [
                    { position: "BOTTOM_AXIS", title: "" },
                    { position: "LEFT_AXIS", title: "金額" },
                  ],
                  domains: [{
                    domain: {
                      sourceRange: { sources: [{ sheetId, startRowIndex: 7, endRowIndex: 12, startColumnIndex: 0, endColumnIndex: 1 }] },
                    },
                  }],
                  series: [
                    {
                      series: {
                        sourceRange: { sources: [{ sheetId, startRowIndex: 7, endRowIndex: 12, startColumnIndex: 2, endColumnIndex: 3 }] },
                      },
                      color: { red: 0.259, green: 0.522, blue: 0.957 },
                    },
                    {
                      series: {
                        sourceRange: { sources: [{ sheetId, startRowIndex: 7, endRowIndex: 12, startColumnIndex: 3, endColumnIndex: 4 }] },
                      },
                      color: { red: 0.204, green: 0.659, blue: 0.325 },
                    },
                  ],
                },
              },
            },
          },
        },
      ],
    },
  });

  console.error("✅ ダッシュボードを初期化しました（KPI数式 + グラフ4つ）");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
