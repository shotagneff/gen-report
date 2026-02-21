/**
 * 提案シート型パイプライン（正本フォーマット：共通＋施策ブロック）
 */

import path from "node:path";
import { chatCompletion, extractJsonBlock } from "./llm.js";
import { writeCsv } from "./csv.js";
import type { CompanyInput, CompanyProfile, ProposalSheet, ProposalRow } from "./types.js";

const TEMPLATE_COMMON_ITEMS = [
  { item: "1システム導入費（初期）", unit: "円", memoHint: "価格帯・条件" },
  { item: "CS/事務の時給（概算）", unit: "円/時", memoHint: "受付・事務" },
  { item: "営業/管理の時給（概算）", unit: "円/時", memoHint: "営業・管理" },
  { item: "駆けつけ/清掃：平均粗利/件", unit: "円/件", memoHint: "変動要因" },
];

const TEMPLATE_INITIATIVES = [
  { name: "AI一次受付（チャット/LINE/フォーム）", items: [{ item: "月の問い合わせ数", unit: "件/月" }, { item: "取りこぼし率（未対応/遅延）", unit: "%" }, { item: "AIで救える割合", unit: "%" }, { item: "問い合わせ→成約率", unit: "%" }, { item: "追加成約1件あたり粗利", unit: "円/件" }] },
  { name: "電話要約→案件登録（受付/事務の工数削減）", items: [{ item: "月の電話/受付件数", unit: "件/月" }, { item: "1件あたり事務時間（現状・受付）", unit: "分/件" }, { item: "削減できる時間（受付）", unit: "分/件" }, { item: "時給（CS/事務）", unit: "円/時" }] },
  { name: "現場レポート自動化（写真+音声→報告書/請求明細）", items: [{ item: "月の案件数（駆けつけ/清掃）", unit: "件/月" }, { item: "1件あたり報告時間（現状・現場）", unit: "分/件" }, { item: "削減できる時間（現場レポ）", unit: "分/件" }, { item: "時給（CS/事務）", unit: "円/時" }] },
  { name: "営業通話分析（勝ちトーク抽出→台本化）", items: [{ item: "月の商談/架電（分析対象）", unit: "件/月" }, { item: "現状成約率", unit: "%" }, { item: "改善幅（pt）", unit: "%" }, { item: "成約1件あたり粗利", unit: "円/件" }] },
  { name: "人材推薦文/マッチング下書き（推薦作業の工数削減）", items: [{ item: "月の推薦/応募処理件数", unit: "件/月" }, { item: "1件あたり作業時間（現状・人材）", unit: "分/件" }, { item: "削減できる時間（人材）", unit: "分/件" }, { item: "時給（採用/人材担当）", unit: "円/時" }] },
  { name: "見積提示の最適化（値引き減/単価UP）", items: [{ item: "月の成約件数（対象）", unit: "件/月" }, { item: "平均粗利/件", unit: "円/件" }, { item: "粗利改善率", unit: "%" }] },
];

export async function buildProposalSheet(
  profile: CompanyProfile,
  input: CompanyInput,
): Promise<ProposalSheet> {
  const systemPrompt = `あなたはAI活用コンサルタントです。
クライアント向けの「考えられる施策」提案シートを、指定のJSON形式だけで出力してください。
共通ブロックは4行（項目・値・単位・メモ）。施策は3〜6件、企業に合うものを選び、各施策は項目ごとに item, value, unit, memo を1行ずつ。数値は想定・目安でよい。出力は指定スキーマのJSONのみ。`;

  const templateText = [
    "【共通ブロックの項目】",
    TEMPLATE_COMMON_ITEMS.map((r) => `- ${r.item}（単位: ${r.unit}）`).join("\n"),
    "",
    "【施策テンプレ一覧】",
    ...TEMPLATE_INITIATIVES.map((s) => `${s.name}: ${s.items.map((i) => i.item).join(", ")}`),
  ].join("\n");

  const userPrompt = `【企業プロファイル】
企業名: ${profile.companyName}
業種: ${profile.industry}
事業概要: ${profile.businessSummary}
キーワード: ${profile.keywords.join(", ")}
注目領域: ${profile.focusAreas.join(", ") || "指定なし"}
${profile.deepSearchSummary ? `ディープサーチ要約: ${profile.deepSearchSummary}` : ""}
${input.memo ? `補足: ${input.memo}` : ""}

${templateText}

上記に基づき、提案シートを以下のJSON形式のみで出力してください。
{
  "companyName": "企業名",
  "title": "考えられる施策",
  "common": {
    "rows": [
      { "item": "項目名", "value": "値", "unit": "単位", "memo": "メモ" }
    ]
  },
  "initiatives": [
    { "name": "施策名", "rows": [ { "item": "項目名", "value": "値", "unit": "単位", "memo": "メモ" } ] }
  ]
}
共通の rows は4件。initiatives は3〜6件。`;

  const content = await chatCompletion([
    { role: "system", content: systemPrompt },
    { role: "user", content: userPrompt },
  ]);

  const parsed = extractJsonBlock(content) as Record<string, unknown>;
  const companyName = String(parsed.companyName ?? profile.companyName);
  const title = typeof parsed.title === "string" ? parsed.title : "考えられる施策";

  const common = parsed.common as { rows?: unknown[] } | undefined;
  const commonRows: ProposalRow[] = Array.isArray(common?.rows)
    ? (common.rows as Record<string, unknown>[]).map((r) => ({
        item: String(r.item ?? ""),
        value: String(r.value ?? ""),
        unit: String(r.unit ?? ""),
        memo: String(r.memo ?? ""),
      }))
    : TEMPLATE_COMMON_ITEMS.map((t) => ({
        item: t.item,
        value: "要入力",
        unit: t.unit,
        memo: t.memoHint,
      }));

  const initiativesRaw = Array.isArray(parsed.initiatives) ? parsed.initiatives : [];
  const initiatives = initiativesRaw.map((inv: Record<string, unknown>) => {
    const name = String(inv.name ?? "");
    const rows = Array.isArray(inv.rows)
      ? (inv.rows as Record<string, unknown>[]).map((r) => ({
          item: String(r.item ?? ""),
          value: String(r.value ?? ""),
          unit: String(r.unit ?? ""),
          memo: String(r.memo ?? ""),
        }))
      : [];
    return { name, rows };
  });

  return {
    companyName,
    title,
    common: { rows: commonRows },
    initiatives,
  };
}

const PROPOSAL_CSV_HEADERS = ["ブロック種別", "施策名", "項目", "値", "単位", "メモ"] as const;

export function writeProposalCsv(
  sheet: ProposalSheet,
  outDir: string,
  filePrefix: string,
): string {
  const outPath = path.join(outDir, `${filePrefix}_proposal_sheet.csv`);
  const rows: string[][] = [];
  for (const r of sheet.common.rows) {
    rows.push(["共通", "", r.item, r.value, r.unit, r.memo]);
  }
  for (const inv of sheet.initiatives) {
    for (const r of inv.rows) {
      rows.push(["施策", inv.name, r.item, r.value, r.unit, r.memo]);
    }
  }
  writeCsv([...PROPOSAL_CSV_HEADERS], rows, outPath);
  return outPath;
}

export function safeFilePrefix(companyName: string): string {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(-2);
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const slug = companyName
    .replace(/\s+/g, "_")
    .replace(/[^\w\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\-_]/g, "")
    .slice(0, 40) || "company";
  return `${yy}_${mm}_${dd}_${slug}御中_AIエージェント活用レポート`;
}
