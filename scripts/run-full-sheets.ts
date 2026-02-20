#!/usr/bin/env npx tsx
/**
 * 4タブ＋前提条件・承認を一括生成。
 * GOOGLE_APPLICATION_CREDENTIALS が設定されていればスプレッドシートにドカンと出力。
 * 未設定なら 5つの CSV を data/out に出力。
 */

import "dotenv/config";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { CompanyInput } from "../src/types.js";
import { resolveInput, buildProfile, extractContactInfo } from "../src/pipeline.js";
import { updateTracking } from "../src/tracking-sheet.js";
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

function parseArgs(argv: string[]): { input: string; outDir: string; sheets: boolean } {
  let input = "";
  let outDir = path.join(repoRoot, "data", "out");
  let sheets = true;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--input" && argv[i + 1]) input = argv[++i];
    else if (argv[i] === "--out-dir" && argv[i + 1]) outDir = argv[++i];
    else if (argv[i] === "--csv-only") sheets = false;
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/run-full-sheets.ts --input <path> [options]

  4タブ（考えられる施策・費用対効果・パッケージ・ロードマップ）＋前提条件・承認を生成。
  認証あり: スプレッドシートに一括出力。認証なし: 5つのCSVを --out-dir に出力。

  --input PATH     Company input JSON (required)
  --out-dir DIR    CSV 出力先 (default: data/out)
  --csv-only       Google Sheets に書き込まず CSV のみ出力
`);
      process.exit(0);
    }
  }
  return { input, outDir, sheets };
}

async function main(): Promise<void> {
  loadDotenv();
  const { input: inputPath, outDir, sheets } = parseArgs(process.argv.slice(2));

  if (!inputPath) {
    console.error("Error: --input is required. Example: --input inputs/example_company.json");
    process.exit(2);
  }

  const absInput = path.isAbsolute(inputPath) ? inputPath : path.join(repoRoot, inputPath);
  if (!fs.existsSync(absInput)) {
    console.error("Error: input file not found:", absInput);
    process.exit(2);
  }

  const raw = JSON.parse(fs.readFileSync(absInput, "utf8")) as CompanyInput;

  console.error("Resolving input...");
  const input = await resolveInput(raw);

  console.error("Building company profile...");
  const profile = await buildProfile(input);

  console.error("Generating 考えられる施策...");
  const proposal = await buildProposalSheet(profile, input);

  console.error("Generating ロードマップ...");
  const roadmap = buildRoadmapRows();

  console.error("Generating 費用対効果 & パッケージ...");
  const { costEffect, packages } = await buildCostEffectAndPackages(proposal);

  const data: FullSpreadsheetData = {
    proposal,
    costEffect,
    packages,
    roadmap,
  };

  const prefix = safeFilePrefix(proposal.companyName);
  const outDirResolved = path.resolve(repoRoot, outDir);

  const hasCreds = !!process.env.GOOGLE_APPLICATION_CREDENTIALS && sheets;

  if (hasCreds) {
    try {
      const url = await writeToGoogleSheets(`${proposal.companyName}御中_AIエージェント活用レポート`, data);
      console.error("Created spreadsheet:", url);
      console.log(JSON.stringify({ spreadsheetUrl: url, tabs: 4 }, null, 2));

      // 営業管理リストに追記
      const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
      const keyPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
      if (folderId && keyPath) {
        try {
          const siteUrl = raw.officialInfo.type === "url" ? raw.officialInfo.value : "";
          const { address, phone } = await extractContactInfo(input.officialInfo.value);
          await updateTracking({ companyName: proposal.companyName, siteUrl, address, phone, reportUrl: url, folderId, credentialsPath: keyPath });
          console.error("✅ 営業リストに追記しました");
        } catch (e) {
          console.error("営業リストへの追記に失敗しました（レポートは生成済み）:", e);
        }
      }
    } catch (e) {
      console.error("Google Sheets export failed, falling back to CSV:", e);
      const paths = writeToCsvFallback(data, outDirResolved, prefix, (h, r, p) =>
        writeCsv(h, r, p),
      );
      console.error("Saved CSV:", paths.join(", "));
      console.log(JSON.stringify({ csvPaths: paths }, null, 2));
    }
  } else {
    const paths = writeToCsvFallback(data, outDirResolved, prefix, (h, r, p) =>
      writeCsv(h, r, p),
    );
    console.error("Saved CSV (set GOOGLE_APPLICATION_CREDENTIALS for Sheets):", paths.join(", "));
    console.log(JSON.stringify({ csvPaths: paths }, null, 2));
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
