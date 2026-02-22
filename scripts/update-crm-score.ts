#!/usr/bin/env npx tsx
/**
 * リード管理CRMのスコアリング列（I〜N）を更新するスクリプト。
 * Usage: npx tsx scripts/update-crm-score.ts --company-name "株式会社〇〇" --score 72 --rank A --contact-type "フォーム" --memo "反応良好" --action "24h以内に電話"
 */

import "dotenv/config";
import { getCRMConnection, findCompanyRow, formatDateYYMMDD } from "../src/crm-common.js";

// ランクに応じたステータス・パイプライン自動更新マッピング
const RANK_STATUS_MAP: Record<string, string> = {
  A: "Aランク対応中",
  B: "ナーチャリング中",
  C: "3ヶ月後フォロー",
};

const RANK_PIPELINE_MAP: Record<string, string> = {
  A: "商談",
  B: "アプローチ中",
  C: "リード",
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
  --no-status-update      ランクに応じたステータス/パイプライン自動更新をスキップ
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

  const conn = await getCRMConnection();
  const tabTitle = conn.tabs.entries().next().value?.[1].title ?? "リスト";

  const match = await findCompanyRow(conn.sheets, conn.spreadsheetId, tabTitle, companyName);
  if (!match) {
    console.error(`Error: 「${companyName}」に一致する会社が見つかりません`);
    process.exit(1);
  }

  const matchedName = match.rowData[1];
  const rowNum = match.rowIndex;

  // スコアリング日を自動設定
  const scoringDate = formatDateYYMMDD();

  // I〜N列を一括更新
  const updateValues = [score, rank, scoringDate, contactType, memo, action];
  await conn.sheets.spreadsheets.values.update({
    spreadsheetId: conn.spreadsheetId,
    range: `'${tabTitle}'!I${rowNum}:N${rowNum}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [updateValues] },
  });

  console.error(`✅ 「${matchedName}」のスコアリングデータを更新しました（行: ${rowNum}）`);
  console.error(`   スコア: ${score} / ランク: ${rank} / 接触経路: ${contactType}`);

  // ランクに応じてステータスとパイプラインを自動更新
  if (updateStatus && rank) {
    const updates: Array<{ range: string; values: string[][] }> = [];

    if (RANK_STATUS_MAP[rank]) {
      updates.push({ range: `'${tabTitle}'!H${rowNum}`, values: [[RANK_STATUS_MAP[rank]]] });
      console.error(`   ステータスを「${RANK_STATUS_MAP[rank]}」に更新しました`);
    }
    if (RANK_PIPELINE_MAP[rank]) {
      updates.push({ range: `'${tabTitle}'!O${rowNum}`, values: [[RANK_PIPELINE_MAP[rank]]] });
      console.error(`   パイプラインを「${RANK_PIPELINE_MAP[rank]}」に更新しました`);
    }

    if (updates.length > 0) {
      await conn.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: conn.spreadsheetId,
        requestBody: { valueInputOption: "USER_ENTERED", data: updates },
      });
    }
  }

  // アクティビティ記録
  try {
    const { appendActivityRow } = await import("../src/tracking-sheet.js");
    await appendActivityRow(conn.sheets, conn.spreadsheetId, {
      companyName: matchedName,
      activityType: "スコアリング更新",
      content: `スコア: ${score} / ランク: ${rank} / 経路: ${contactType}${memo ? ` / メモ: ${memo}` : ""}`,
      result: action,
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
