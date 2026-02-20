/**
 * 共通パイプライン：入力解決・企業プロファイル生成
 * レポート用・提案シート用の両方で利用
 */

import { chatCompletion, extractJsonBlock } from "./llm.js";
import { fetchUrlAsText } from "./fetch-url.js";
import type { CompanyInput, CompanyProfile } from "./types.js";

/** Agent 1: 入力の正規化と URL → テキスト取得 */
export async function resolveInput(raw: CompanyInput): Promise<CompanyInput> {
  if (!raw.companyName?.trim()) {
    throw new Error("企業名がありません");
  }
  if (!raw.officialInfo?.value?.trim()) {
    throw new Error("公式情報（URL または テキスト）がありません");
  }

  const officialInfo = { ...raw.officialInfo };
  if (officialInfo.type === "url") {
    const url = officialInfo.value.trim();
    try {
      const text = await fetchUrlAsText(url);
      officialInfo.type = "text";
      officialInfo.value = text.slice(0, 8000);
    } catch (e) {
      throw new Error(`URLの取得に失敗しました: ${url} - ${String(e)}`);
    }
  }

  return {
    ...raw,
    companyName: raw.companyName.trim(),
    officialInfo,
  };
}

/** Agent 2: 企業プロファイルの生成 */
export async function buildProfile(input: CompanyInput): Promise<CompanyProfile> {
  const systemPrompt = `あなたは企業情報を構造化するアシスタントです。
渡された公式情報（とオプション）から、業種・事業概要・キーワードを抽出し、指定のJSON形式のみを返してください。
推測は最小限にし、書かれている事実を優先します。業種が不明な場合は「その他」とします。`;

  const parts: string[] = [
    `企業名: ${input.companyName}`,
    "",
    "【公式情報】",
    input.officialInfo.value.slice(0, 6000),
  ];
  if (input.industry) parts.push("", "【指定業種】", input.industry);
  if (input.businessSummary) parts.push("", "【指定事業概要】", input.businessSummary);
  if (input.focusAreas?.length) {
    parts.push("", "【注目領域】", input.focusAreas.join(", "));
  }
  if (input.deepSearchResult) {
    parts.push("", "【ディープサーチ結果（参考）】", input.deepSearchResult.slice(0, 3000));
  }

  const userPrompt = `${parts.join("\n")}

以下のJSON形式のみを返してください（他に説明は不要）。キー名は必ず英語にすること。
{
  "companyName": "企業名",
  "industry": "業種",
  "businessSummary": "事業概要（1〜2文）",
  "keywords": ["キーワード1", "キーワード2"],
  "focusAreas": ["注目領域1"],
  "deepSearchSummary": "ディープサーチの要約（あれば1〜2文）"
}`;

  const content = await chatCompletion([
    { role: "system", content: systemPrompt },
    { role: "user", content: userPrompt },
  ]);

  const parsed = extractJsonBlock(content) as Record<string, unknown>;
  return {
    companyName: String(parsed.companyName ?? input.companyName),
    industry: String(parsed.industry ?? "その他"),
    businessSummary: String(parsed.businessSummary ?? ""),
    keywords: Array.isArray(parsed.keywords)
      ? (parsed.keywords as string[]).map(String)
      : [],
    focusAreas: Array.isArray(parsed.focusAreas)
      ? (parsed.focusAreas as string[]).map(String)
      : input.focusAreas ?? [],
    deepSearchSummary:
      typeof parsed.deepSearchSummary === "string"
        ? parsed.deepSearchSummary
        : undefined,
  };
}

/** Agent: 住所・電話番号を抽出する */
export async function extractContactInfo(
  text: string,
): Promise<{ address: string; phone: string }> {
  const userPrompt = `以下のテキストから会社の住所と代表電話番号を1つずつ抽出してください。
見つからない場合は空文字にしてください。
JSON形式のみを返してください（説明不要）。

${text.slice(0, 4000)}

{"address":"...","phone":"..."}`;

  try {
    const content = await chatCompletion([
      { role: "user", content: userPrompt },
    ]);
    const parsed = extractJsonBlock(content) as Record<string, unknown>;
    return {
      address: typeof parsed.address === "string" ? parsed.address : "",
      phone: typeof parsed.phone === "string" ? parsed.phone : "",
    };
  } catch {
    return { address: "", phone: "" };
  }
}
