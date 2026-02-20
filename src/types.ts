/**
 * 企業AI活用レポート用の型定義（agent-design.md と対応）
 */

export interface CompanyInput {
  companyName: string;
  officialInfo: { type: "url" | "text"; value: string };
  industry?: string;
  businessSummary?: string;
  focusAreas?: string[];
  existingAiDx?: string;
  deepSearchResult?: string;
  memo?: string;
}

export interface CompanyProfile {
  companyName: string;
  industry: string;
  businessSummary: string;
  keywords: string[];
  focusAreas: string[];
  deepSearchSummary?: string;
}

export interface UtilizationRow {
  companyName: string;
  no: number;
  scene: string;
  method: string;
  effectQualitative: string;
  effectQuantitative?: string;
  priority: "高" | "中" | "低";
  note?: string;
}

/** 出力用：サマリ1行 */
export interface SummaryRow {
  companyName: string;
  industry: string;
  businessSummary: string;
  officialUrl: string;
  generatedAt: string;
  focusAreasMemo: string;
}

/** 提案シート型：共通・施策の1行 */
export interface ProposalRow {
  item: string;
  value: string;
  unit: string;
  memo: string;
}

export interface CommonBlock {
  rows: ProposalRow[];
}

export interface InitiativeBlock {
  name: string;
  rows: ProposalRow[];
}

export interface ProposalSheet {
  companyName: string;
  title?: string;
  common: CommonBlock;
  initiatives: InitiativeBlock[];
}

// --- 4タブ構成（費用対効果・パッケージ・ロードマップ・前提条件・承認）---

/** 費用対効果など：施策ごとの効果・回収・ROI 1行 */
export interface CostEffectRow {
  measure: string;       // 施策1: AI一次受付 ...
  effectType: string;   // 取りこぼし削減(粗利増) など
  monthlyEffectYen: number;
  yearlyEffectYen: number;
  paybackDays: number;
  yearlyRoiPercent: number;
  mainDriver: string;
  kpi: string;
  memo: string;
}

/** パッケージ提案 1行 */
export interface PackageRow {
  packageName: string;   // A: Sales Boost
  content: string;      // 施策4+施策6
  benefit: string;     // 想定メリット
  systemCount: number;
  initialCostYen: number;
  monthlyEffectYen: number;
  paybackMonths: number;
  memo: string;
}

/** ロードマップ 1行（工程 + W1..W8） */
export interface RoadmapRow {
  phase: string;        // 0. Kickoff
  content: string;      // 内容
  w1: string;           // '' or '1' or '●'
  w2: string;
  w3: string;
  w4: string;
  w5: string;
  w6: string;
  w7: string;
  w8: string;
}

/** 前提条件・承認 1行 */
export interface PremiseRow {
  premise: string;      // 前提条件
  category: string;     // 考えられる施策用 / 費用対効果用
  status: string;       // 未確認 / 確認済み
  memo: string;
}
