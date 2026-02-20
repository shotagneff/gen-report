#!/usr/bin/env npx tsx
/**
 * 提案シート型の生成物を出力（正本フォーマット：共通＋施策ブロック）
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { CompanyInput } from "../src/types.js";
import { resolveInput, buildProfile } from "../src/pipeline.js";
import {
  buildProposalSheet,
  writeProposalCsv,
  safeFilePrefix,
} from "../src/pipeline-proposal.js";

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

function parseArgs(argv: string[]): { input: string; outDir: string } {
  let input = "";
  let outDir = path.join(repoRoot, "data", "out");
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--input" && argv[i + 1]) input = argv[++i];
    else if (argv[i] === "--out-dir" && argv[i + 1]) outDir = argv[++i];
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log("Usage: npx tsx scripts/run-proposal.ts --input <path> [--out-dir <dir>]");
      process.exit(0);
    }
  }
  return { input, outDir };
}

async function main(): Promise<void> {
  loadDotenv();
  const { input: inputPath, outDir } = parseArgs(process.argv.slice(2));

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

  console.error("Generating proposal sheet (common + initiatives)...");
  const sheet = await buildProposalSheet(profile, input);

  const prefix = safeFilePrefix(sheet.companyName);
  const outPath = writeProposalCsv(sheet, path.resolve(repoRoot, outDir), prefix);

  console.error("Saved:", outPath);
  console.log(JSON.stringify({ proposalSheetPath: outPath, initiativesCount: sheet.initiatives.length }, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
