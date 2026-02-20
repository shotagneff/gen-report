/**
 * ディープクロール：ホームページ＋関連下層ページを取得してテキストを結合する
 * 追加ライブラリ不要（Node.js fetch + regex）
 */

const RELEVANT_PATTERNS = [
  /about/i, /company/i, /corporate/i, /overview/i,
  /service/i, /solution/i, /business/i, /product/i,
  /news/i, /press/i, /case/i, /recruit/i,
  /会社/i, /事業/i, /サービス/i, /ソリューション/i,
  /実績/i, /事例/i, /採用/i, /概要/i,
];

const SKIP_EXTENSIONS = /\.(jpg|jpeg|png|gif|svg|webp|pdf|zip|css|js|woff|ttf|ico|xml|json)(\?.*)?$/i;

function extractInternalLinks(html: string, baseUrl: string): string[] {
  const base = new URL(baseUrl);
  const seen = new Set<string>();
  const links: string[] = [];

  const hrefRegex = /href=["']([^"'#?][^"']*?)["']/gi;
  let m: RegExpExecArray | null;
  while ((m = hrefRegex.exec(html)) !== null) {
    const raw = m[1].trim();
    if (!raw || SKIP_EXTENSIONS.test(raw)) continue;
    try {
      const abs = new URL(raw, base).href;
      const u = new URL(abs);
      if (u.hostname !== base.hostname) continue;
      if (seen.has(abs)) continue;
      seen.add(abs);
      links.push(abs);
    } catch {
      // ignore invalid URLs
    }
  }
  return links;
}

function scoreLink(url: string): number {
  let score = 0;
  for (const p of RELEVANT_PATTERNS) {
    if (p.test(url)) score += 1;
  }
  // 深すぎるパスは優先度下げ
  const depth = (url.match(/\//g) ?? []).length;
  if (depth > 5) score -= 1;
  return score;
}

async function fetchHtml(url: string): Promise<string> {
  const res = await fetch(url, {
    headers: { "User-Agent": "CompanyAIReport/1.0 (research)" },
    signal: AbortSignal.timeout(8000),
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.text();
}

function htmlToText(html: string): string {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * URL からホームページ＋関連下層ページをクロールして結合テキストを返す
 * @param url 企業の公式サイトURL
 * @param maxSubPages 下層ページの最大取得数（デフォルト4）
 */
export async function deepCrawl(url: string, maxSubPages = 4): Promise<string> {
  const parts: string[] = [];

  // ホームページ取得
  let homeHtml: string;
  try {
    process.stderr.write(`  [deep-crawl] ホームページ取得中: ${url}\n`);
    homeHtml = await fetchHtml(url);
    parts.push(`=== ホームページ ===\n${htmlToText(homeHtml).slice(0, 4000)}`);
  } catch (e) {
    throw new Error(`ホームページの取得に失敗しました: ${url} - ${String(e)}`);
  }

  // 下層ページのリンク抽出 → スコア順にソート
  const links = extractInternalLinks(homeHtml, url);
  const scored = links
    .map((link) => ({ link, score: scoreLink(link) }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .map((x) => x.link)
    .slice(0, maxSubPages * 2); // 候補を多めに取り、失敗を考慮

  // 下層ページを順次取得（最大 maxSubPages 件）
  let fetched = 0;
  for (const subUrl of scored) {
    if (fetched >= maxSubPages) break;
    try {
      process.stderr.write(`  [deep-crawl] サブページ取得中: ${subUrl}\n`);
      const html = await fetchHtml(subUrl);
      const text = htmlToText(html).slice(0, 2000);
      parts.push(`=== ${subUrl} ===\n${text}`);
      fetched++;
    } catch {
      // 取得失敗は無視して次へ
    }
  }

  process.stderr.write(`  [deep-crawl] 完了: ホームページ + ${fetched}サブページ\n`);
  return parts.join("\n\n").slice(0, 14000);
}
