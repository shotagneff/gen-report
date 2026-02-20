/**
 * 4タブ＋前提条件・承認のデータ組み立て
 * 考えられる施策は既存の ProposalSheet。ここでは 費用対効果・パッケージ・ロードマップ・前提条件 を生成する。
 */

import { chatCompletion, extractJsonBlock } from "./llm.js";
import type {
  ProposalSheet,
  CostEffectRow,
  PackageRow,
  RoadmapRow,
  PremiseRow,
} from "./types.js";

/** 前提条件・承認：共通＋施策名から生成（承認ステータスは未確認） */
export function buildPremiseRows(sheet: ProposalSheet): PremiseRow[] {
  const rows: PremiseRow[] = [];
  for (const r of sheet.common.rows) {
    rows.push({
      premise: `${r.item}: ${r.value} ${r.unit}`,
      category: "考えられる施策用",
      status: "未確認",
      memo: r.memo || "",
    });
  }
  for (const inv of sheet.initiatives) {
    rows.push({
      premise: `施策: ${inv.name}`,
      category: "考えられる施策用",
      status: "未確認",
      memo: "採用可否を確認",
    });
  }
  return rows;
}

/** ロードマップ：8週間 PoC→パイロット→本番（固定テンプレ） */
export function buildRoadmapRows(): RoadmapRow[] {
  return [
    { phase: "0. Kickoff", content: "目的/KPI合意、現状フロー確認、対象施策の確定", w1: "●", w2: "", w3: "", w4: "", w5: "", w6: "", w7: "", w8: "" },
    { phase: "1. データ/運用設計", content: "入力項目、権限、ログ、テンプレ(受付/報告/台本)", w1: "", w2: "●", w3: "", w4: "", w5: "", w6: "", w7: "", w8: "" },
    { phase: "2. 実装", content: "AIフロー実装、連携(フォーム/LINE/CRM/予約/スプレッドシート等)", w1: "", w2: "●", w3: "●", w4: "●", w5: "", w6: "", w7: "", w8: "" },
    { phase: "3. PoC", content: "限定チャネルで検証(1拠点/1チーム/1媒体)", w1: "", w2: "", w3: "", w4: "●", w5: "●", w6: "", w7: "", w8: "" },
    { phase: "4. パイロット", content: "運用を広げる(週次で改善、FAQ/台本のチューニング)", w1: "", w2: "", w3: "", w4: "", w5: "●", w6: "●", w7: "", w8: "" },
    { phase: "5. 本番展開", content: "全チャネル/全員へ展開、教育、運用ルール定着", w1: "", w2: "", w3: "", w4: "", w5: "", w6: "", w7: "●", w8: "" },
    { phase: "6. KPI レビュー", content: "効果測定→次施策の優先順位決定(四半期計画へ)", w1: "", w2: "", w3: "", w4: "", w5: "", w6: "", w7: "", w8: "●" },
  ];
}

/** 費用対効果＋パッケージをLLMで生成 */
export async function buildCostEffectAndPackages(
  sheet: ProposalSheet,
): Promise<{ costEffect: CostEffectRow[]; packages: PackageRow[] }> {
  const systemPrompt = `あなたはAI活用コンサルタントです。
「考えられる施策」の共通パラメータと各施策の項目・値をもとに、以下2つを指定のJSON形式のみで出力してください。
1) 施策ごとの効果・回収・ROI: 各施策について 効果タイプ、月次効果(円)、年次効果(円)、回収期間(日)、年次ROI(%)、主要ドライバー、測定KPI、メモ。数値は想定でよい。
2) パッケージ提案: 3案。例) A=営業系施策2つ、B=事務・現場系3つ、C=全施策。各パッケージで 内容(導入システム)、想定メリット、システム数、初期費用(円)、月次効果(円)、回収期間(月)、メモ。出力は指定スキーマのJSONのみ。`;

  const initiativesText = sheet.initiatives
    .map((inv) => `${inv.name}: ${inv.rows.map((r) => `${r.item}=${r.value}${r.unit}`).join(", ")}`)
    .join("\n");
  const commonText = sheet.common.rows.map((r) => `${r.item}=${r.value} ${r.unit}`).join(", ");

  const userPrompt = `【企業名】${sheet.companyName}
【共通】${commonText}
【施策】
${initiativesText}

上記に基づき、以下JSONのみ出力。
{
  "costEffect": [
    {
      "measure": "施策1: AI一次受付(チャット/LINE/フォーム)",
      "effectType": "取りこぼし削減(粗利増)",
      "monthlyEffectYen": 38000,
      "yearlyEffectYen": 456000,
      "paybackDays": 13,
      "yearlyRoiPercent": 91.2,
      "mainDriver": "問い合わせ数×取りこぼし率×回収率×成約率×粗利",
      "kpi": "未対応率、予約完了率、成約率",
      "memo": "夜間・繁忙で効果が出やすい"
    }
  ],
  "packages": [
    {
      "packageName": "A: Sales Boost",
      "content": "施策4+施策6",
      "benefit": "成約率/粗利の底上げ",
      "systemCount": 2,
      "initialCostYen": 1000000,
      "monthlyEffectYen": 267000,
      "paybackMonths": 3.7,
      "memo": "営業代行が主力なら最適"
    }
  ]
}
costEffect は施策数分の要素。packages は3件。`;

  const content = await chatCompletion([
    { role: "system", content: systemPrompt },
    { role: "user", content: userPrompt },
  ]);

  const parsed = extractJsonBlock(content) as {
    costEffect?: unknown[];
    packages?: unknown[];
  };

  const costEffect: CostEffectRow[] = (Array.isArray(parsed.costEffect) ? parsed.costEffect : []).map(
    (r: Record<string, unknown>) => ({
      measure: String(r.measure ?? ""),
      effectType: String(r.effectType ?? ""),
      monthlyEffectYen: Number(r.monthlyEffectYen) || 0,
      yearlyEffectYen: Number(r.yearlyEffectYen) || 0,
      paybackDays: Number(r.paybackDays) || 0,
      yearlyRoiPercent: Number(r.yearlyRoiPercent) || 0,
      mainDriver: String(r.mainDriver ?? ""),
      kpi: String(r.kpi ?? ""),
      memo: String(r.memo ?? ""),
    }),
  );

  const packages: PackageRow[] = (Array.isArray(parsed.packages) ? parsed.packages : []).map(
    (r: Record<string, unknown>) => ({
      packageName: String(r.packageName ?? ""),
      content: String(r.content ?? ""),
      benefit: String(r.benefit ?? ""),
      systemCount: Number(r.systemCount) || 0,
      initialCostYen: Number(r.initialCostYen) || 0,
      monthlyEffectYen: Number(r.monthlyEffectYen) || 0,
      paybackMonths: Number(r.paybackMonths) || 0,
      memo: String(r.memo ?? ""),
    }),
  );

  return { costEffect, packages };
}

export type CostEffectRowType = "title" | "tableHeader" | "data" | "empty" | "total";
export interface CostEffectVisualRow { row: (string | number)[]; type: CostEffectRowType; }

/** 費用対効果のビジュアル行を生成（Sheets フォーマット適用用） */
export function costEffectToVisualRows(rows: CostEffectRow[]): CostEffectVisualRow[] {
  const result: CostEffectVisualRow[] = [];
  result.push({ row: ["施策ごとの効果・回収・ROI", ...Array(8).fill("")], type: "title" });
  result.push({ row: ["施策", "効果タイプ", "月次効果（円）", "年次効果（円）", "回収期間（日）", "年次ROI（1年目）", "主要ドライバー（Inputs参照）", "測定KPI", "メモ"], type: "tableHeader" });
  for (const r of rows) {
    result.push({ row: [r.measure, r.effectType, r.monthlyEffectYen, r.yearlyEffectYen, r.paybackDays, r.yearlyRoiPercent, r.mainDriver, r.kpi, r.memo], type: "data" });
  }
  result.push({ row: Array(9).fill(""), type: "empty" });
  const totalM = rows.reduce((s, r) => s + r.monthlyEffectYen, 0);
  const totalY = rows.reduce((s, r) => s + r.yearlyEffectYen, 0);
  const avgP = rows.length ? rows.reduce((s, r) => s + r.paybackDays, 0) / rows.length : 0;
  const avgR = rows.length ? rows.reduce((s, r) => s + r.yearlyRoiPercent, 0) / rows.length : 0;
  result.push({ row: ["合計（参考）", "", totalM, totalY, Math.round(avgP * 10) / 10, Math.round(avgR * 10) / 10, "", "", ""], type: "total" });
  return result;
}

/** 費用対効果に合計行を追加した2次元配列（CSV フォールバック用） */
export function costEffectToSheetRows(rows: CostEffectRow[]): string[][] {
  const header = ["施策", "効果タイプ", "月次効果(円)", "年次効果(円)", "回収期間(日)", "年次ROI(1年目)", "主要ドライバー(Inputs参照)", "測定KPI", "メモ"];
  const data = rows.map((r) => [
    r.measure,
    r.effectType,
    String(r.monthlyEffectYen),
    String(r.yearlyEffectYen),
    String(r.paybackDays),
    String(r.yearlyRoiPercent),
    r.mainDriver,
    r.kpi,
    r.memo,
  ]);
  const totalMonthly = rows.reduce((s, r) => s + r.monthlyEffectYen, 0);
  const totalYearly = rows.reduce((s, r) => s + r.yearlyEffectYen, 0);
  const avgPayback = rows.length ? rows.reduce((s, r) => s + r.paybackDays, 0) / rows.length : 0;
  const avgRoi = rows.length ? rows.reduce((s, r) => s + r.yearlyRoiPercent, 0) / rows.length : 0;
  data.push(["合計(参考)", "", String(totalMonthly), String(totalYearly), String(Math.round(avgPayback * 10) / 10), String(Math.round(avgRoi * 10) / 10), "", "", ""]);
  return [header, ...data];
}

export function packageToSheetRows(rows: PackageRow[]): string[][] {
  const header = ["パッケージ", "内容(導入システム)", "想定メリット", "システム数", "初期費用(円)", "月次効果(円)", "回収期間(月)", "メモ"];
  const data = rows.map((r) => [
    r.packageName,
    r.content,
    r.benefit,
    String(r.systemCount),
    String(r.initialCostYen),
    String(r.monthlyEffectYen),
    String(r.paybackMonths),
    r.memo,
  ]);
  return [header, ...data];
}

export function roadmapToSheetRows(rows: RoadmapRow[]): string[][] {
  const header = ["工程", "内容", "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8"];
  const data = rows.map((r) => [r.phase, r.content, r.w1, r.w2, r.w3, r.w4, r.w5, r.w6, r.w7, r.w8]);
  return [header, ...data];
}

export type RoadmapRowType = "title" | "empty" | "tableHeader" | "data";
export interface RoadmapVisualRow { row: string[]; type: RoadmapRowType; weekMask?: boolean[]; }

/** ロードマップのビジュアル行を生成（Sheets フォーマット適用用） */
export function roadmapToVisualRows(rows: RoadmapRow[]): RoadmapVisualRow[] {
  const result: RoadmapVisualRow[] = [];
  result.push({ row: ["ロードマップ（8週間：PoC→パイロット→本番）", ...Array(9).fill("")], type: "title" });
  result.push({ row: Array(10).fill(""), type: "empty" });
  result.push({ row: ["工程", "内容", "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8"], type: "tableHeader" });
  for (const r of rows) {
    result.push({
      row: [r.phase, r.content, "", "", "", "", "", "", "", ""],
      type: "data",
      weekMask: [r.w1 !== "", r.w2 !== "", r.w3 !== "", r.w4 !== "", r.w5 !== "", r.w6 !== "", r.w7 !== "", r.w8 !== ""],
    });
  }
  return result;
}

export type PackageRowType = "title" | "empty" | "tableHeader" | "data";
export interface PackageVisualRow { row: (string | number)[]; type: PackageRowType; }

/** パッケージのビジュアル行を生成（Sheets フォーマット適用用） */
export function packageToVisualRows(rows: PackageRow[]): PackageVisualRow[] {
  const result: PackageVisualRow[] = [];
  result.push({ row: ["パッケージ提案", ...Array(7).fill("")], type: "title" });
  result.push({ row: Array(8).fill(""), type: "empty" });
  result.push({ row: ["パッケージ", "内容（導入システム）", "想定メリット", "システム数", "初期費用（円）", "月次効果（円）", "回収期間（月）", "メモ"], type: "tableHeader" });
  for (const r of rows) {
    result.push({ row: [r.packageName, r.content, r.benefit, r.systemCount, r.initialCostYen, r.monthlyEffectYen, r.paybackMonths, r.memo], type: "data" });
  }
  return result;
}

export function premiseToSheetRows(rows: PremiseRow[]): string[][] {
  const header = ["前提条件", "区分", "承認ステータス", "メモ"];
  const data = rows.map((r) => [r.premise, r.category, r.status, r.memo]);
  return [header, ...data];
}
