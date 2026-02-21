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
  siteUrl?: string,
  companyName?: string,
): Promise<{ address: string; phone: string }> {
  // まずメインテキストから抽出を試みる
  let combined = text.slice(0, 4000);

  // URLが渡された場合、会社概要ページも取得して情報を補完する
  if (siteUrl) {
    const baseUrl = siteUrl.replace(/\/+$/, "");
    const companyPaths = ["/company", "/about", "/corporate", "/outline"];
    for (const p of companyPaths) {
      try {
        const pageText = await fetchUrlAsText(`${baseUrl}${p}`);
        if (pageText && pageText.length > 100) {
          combined = `${combined}\n\n--- 会社概要ページ ---\n${pageText.slice(0, 4000)}`;
          break;
        }
      } catch {
        // 存在しないパスはスキップ
      }
    }
  }

  let result = await extractContactFromText(combined);

  // サイトから取得できなかった場合、DuckDuckGoで補完
  if ((!result.address || !result.phone) && companyName) {
    try {
      const ddgUrl = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(companyName + " 住所 電話番号")}`;
      const res = await fetch(ddgUrl, {
        headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36" },
      });
      if (res.ok) {
        const html = await res.text();
        const searchText = html
          .replace(/<script[\s\S]*?<\/script>/gi, "")
          .replace(/<style[\s\S]*?<\/style>/gi, "")
          .replace(/<[^>]+>/g, " ")
          .replace(/\s+/g, " ")
          .trim();
        if (searchText.length > 200) {
          const searchResult = await extractContactFromText(
            `会社名: ${companyName}\n\n--- Web検索結果 ---\n${searchText.slice(0, 6000)}`
          );
          if (!result.address && searchResult.address) result.address = searchResult.address;
          if (!result.phone && searchResult.phone) result.phone = searchResult.phone;
        }
      }
    } catch {
      // 検索失敗はスキップ
    }
  }

  return result;
}

async function extractContactFromText(
  text: string,
): Promise<{ address: string; phone: string }> {
  const userPrompt = `以下のテキストから会社の住所と代表電話番号を1つずつ抽出してください。
見つからない場合は空文字にしてください。
JSON形式のみを返してください（説明不要）。

${text}

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
