/**
 * URL からテキストを取得（簡易：HTML タグ除去）
 * 本番では Puppeteer / Playwright / MCP などで JavaScript レンダリングが必要な場合は差し替え
 */

export async function fetchUrlAsText(url: string): Promise<string> {
  const res = await fetch(url, {
    headers: { "User-Agent": "CompanyAIReport/1.0 (research)" },
  });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}: ${url}`);
  }
  const html = await res.text();
  return stripHtml(html);
}

function stripHtml(html: string): string {
  return (
    html
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 12000) // コンテキスト節約
  );
}
