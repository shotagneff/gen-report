#!/usr/bin/env npx tsx
/**
 * インタラクティブCLI：質問に答えるだけでスプレッドシートを生成する
 */

import "dotenv/config";
import * as readline from "node:readline/promises";
import { stdin as input, stdout as output, stderr } from "node:process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { CompanyInput } from "../src/types.js";
import { resolveInput, buildProfile } from "../src/pipeline.js";
import { deepCrawl } from "../src/deep-crawl.js";
import { buildProposalSheet, safeFilePrefix } from "../src/pipeline-proposal.js";
import {
  buildRoadmapRows,
  buildCostEffectAndPackages,
} from "../src/build-four-tabs.js";
import {
  writeToGoogleSheets,
  writeToCsvFallback,
  type FullSpreadsheetData,
} from "../src/sheets-export.js";
import { writeCsv } from "../src/csv.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");

function log(msg: string) {
  stderr.write(msg + "\n");
}

function printHeader() {
  log("");
  log("╔══════════════════════════════════════════════════╗");
  log("║   AI活用レポート ジェネレーター                    ║");
  log("╚══════════════════════════════════════════════════╝");
  log("  会社名とURLを入力するだけでスプレッドシートを自動生成します。");
  log("  ※ Enterだけで省略できる項目もあります。");
  log("");
}

async function prompt(rl: readline.Interface, question: string, required = false): Promise<string> {
  while (true) {
    const answer = (await rl.question(question)).trim();
    if (answer || !required) return answer;
    stderr.write("  ⚠️  この項目は必須です。入力してください。\n");
  }
}

async function main(): Promise<void> {
  printHeader();

  const rl = readline.createInterface({ input, output });

  // ① 会社名
  const companyName = await prompt(rl, "? 会社名を入力してください: ", true);

  // ② URL or テキスト
  log("");
  log("  公式サイトのURL、または会社の説明文を入力してください。");
  log("  URLを入れると自動でホームページを読み込みます。");
  const officialValue = await prompt(rl, "? URL または 説明文: ", true);
  const isUrl = officialValue.startsWith("http://") || officialValue.startsWith("https://");

  // ③ ディープクロール（URLの場合のみ）
  let useDeepCrawl = false;
  if (isUrl) {
    log("");
    const ans = await prompt(rl, "? ホームページの下層ページも自動収集しますか？精度が上がります。[Y/n]: ");
    useDeepCrawl = ans === "" || ans.toLowerCase() === "y";
  }

  // ④ 業種
  log("");
  const industry = await prompt(rl, "? 業種（例: IT・SaaS, 建設, 不動産 ／ スキップはEnter）: ");

  // ⑤ 注力部門
  log("");
  log("  AIを活用したい部門をカンマ区切りで入力してください。");
  const focusAreasRaw = await prompt(rl, "? 注力部門（例: 営業,現場,経理 ／ スキップはEnter）: ");
  const focusAreas = focusAreasRaw
    ? focusAreasRaw.split(",").map((s) => s.trim()).filter(Boolean)
    : [];

  // ⑥ 補足メモ
  log("");
  const memo = await prompt(rl, "? 補足メモ（担当者コメント等 ／ スキップはEnter）: ");

  rl.close();

  log("");
  log("────────────────────────────────────────────────────");
  log(`  会社名    : ${companyName}`);
  log(`  情報元    : ${officialValue.slice(0, 60)}${officialValue.length > 60 ? "..." : ""}`);
  if (isUrl) log(`  深掘り    : ${useDeepCrawl ? "あり（下層ページも取得）" : "なし（トップページのみ）"}`);
  if (industry) log(`  業種      : ${industry}`);
  if (focusAreas.length) log(`  注力部門  : ${focusAreas.join(", ")}`);
  if (memo) log(`  メモ      : ${memo}`);
  log("────────────────────────────────────────────────────");
  log("");

  // officialInfo 組み立て
  let officialInfo: CompanyInput["officialInfo"];
  if (isUrl && useDeepCrawl) {
    log("ディープクロール中...");
    const crawledText = await deepCrawl(officialValue);
    officialInfo = { type: "text", value: crawledText };
  } else {
    officialInfo = { type: isUrl ? "url" : "text", value: officialValue };
  }

  const raw: CompanyInput = {
    companyName,
    officialInfo,
    ...(industry && { industry }),
    ...(focusAreas.length && { focusAreas }),
    ...(memo && { memo }),
  };

  log("入力情報を解析中...");
  const resolvedInput = await resolveInput(raw);

  log("企業プロファイルを生成中...");
  const profile = await buildProfile(resolvedInput);

  log("考えられる施策を生成中...");
  const proposal = await buildProposalSheet(profile, resolvedInput);

  log("ロードマップを生成中...");
  const roadmap = buildRoadmapRows();

  log("費用対効果・パッケージを生成中...");
  const { costEffect, packages } = await buildCostEffectAndPackages(proposal);

  const data: FullSpreadsheetData = { proposal, costEffect, packages, roadmap };
  const prefix = safeFilePrefix(proposal.companyName);
  const outDir = path.join(repoRoot, "data", "out");

  const hasCreds = !!process.env.GOOGLE_APPLICATION_CREDENTIALS;

  if (hasCreds) {
    try {
      log("スプレッドシートを作成中...");
      const url = await writeToGoogleSheets(`${proposal.companyName}御中_AIエージェント活用レポート`, data);
      log("");
      log("✅ 完了！スプレッドシートのURLはこちら:");
      log(`   ${url}`);
      log("");
      console.log(JSON.stringify({ spreadsheetUrl: url, tabs: 4 }, null, 2));
    } catch (e) {
      log(`Google Sheets エクスポート失敗: ${String(e)}`);
      log("CSVにフォールバックします...");
      const paths = writeToCsvFallback(data, outDir, prefix, (h, r, p) => writeCsv(h, r, p));
      log(`✅ CSV出力完了: ${paths.join(", ")}`);
      console.log(JSON.stringify({ csvPaths: paths }, null, 2));
    }
  } else {
    const paths = writeToCsvFallback(data, outDir, prefix, (h, r, p) => writeCsv(h, r, p));
    log("");
    log("✅ CSV出力完了（Google Sheets を使う場合は GOOGLE_APPLICATION_CREDENTIALS を設定してください）:");
    paths.forEach((p) => log(`   ${p}`));
    log("");
    console.log(JSON.stringify({ csvPaths: paths }, null, 2));
  }
}

main().catch((err) => {
  stderr.write(`\nエラー: ${String(err)}\n`);
  process.exit(1);
});
