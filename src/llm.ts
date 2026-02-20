/**
 * LLM 呼び出し（OpenAI API 互換）
 * OPENAI_API_KEY + OPENAI_BASE_URL または ANTHROPIC_API_KEY などに差し替え可能
 */

const DEFAULT_BASE = "https://api.openai.com/v1";
const DEFAULT_MODEL = "gpt-4o-mini";

export interface LlmConfig {
  apiKey: string;
  baseUrl?: string;
  model?: string;
}

function getConfig(): LlmConfig {
  const apiKey = process.env.OPENAI_API_KEY ?? process.env.ANTHROPIC_API_KEY ?? "";
  const baseUrl = process.env.OPENAI_BASE_URL ?? process.env.ANTHROPIC_BASE_URL ?? DEFAULT_BASE;
  const model = process.env.OPENAI_MODEL ?? process.env.ANTHROPIC_MODEL ?? DEFAULT_MODEL;
  return { apiKey, baseUrl, model };
}

/**
 * Chat completion を呼び、助手のテキストを1つ返す。
 * OpenAI: /v1/chat/completions, Anthropic: 適宜 baseUrl を変更
 */
export async function chatCompletion(
  messages: Array<{ role: "system" | "user" | "assistant"; content: string }>,
  config?: Partial<LlmConfig>,
): Promise<string> {
  const cfg = { ...getConfig(), ...config };
  if (!cfg.apiKey) {
    throw new Error("OPENAI_API_KEY or ANTHROPIC_API_KEY must be set");
  }

  const url = cfg.baseUrl!.replace(/\/$/, "") + "/chat/completions";
  const body: Record<string, unknown> = {
    model: cfg.model ?? DEFAULT_MODEL,
    messages,
    max_tokens: 4096,
  };

  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${cfg.apiKey}`,
    },
    body: JSON.stringify(body),
  });

  const text = await res.text();
  if (!res.ok) {
    throw new Error(`LLM API ${res.status}: ${text.slice(0, 2000)}`);
  }

  const data = JSON.parse(text) as { choices?: Array<{ message?: { content?: string } }> };
  const content = data.choices?.[0]?.message?.content?.trim();
  if (!content) {
    throw new Error("LLM returned empty content");
  }
  return content;
}

/**
 * 返り値のテキストから JSON ブロックを1つ抽出してパースする
 */
export function extractJsonBlock(text: string): unknown {
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) {
    throw new Error("No JSON object found in response");
  }
  return JSON.parse(match[0]) as unknown;
}
