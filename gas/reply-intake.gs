/**
 * reply-intake: 返信メール自動検知 → AI解析 → CRM自動登録
 *
 * 設定手順:
 * 1. CRMスプレッドシート → 拡張機能 → Apps Script
 * 2. このコードを貼り付け
 * 3. スクリプトプロパティに以下を設定:
 *    - ANTHROPIC_API_KEY: Claude APIキー
 *    - SHEET_TAB_NAME: タブ名（デフォルト: リスト）
 * 4. トリガー設定: checkReplyEmails を5分おきに実行
 */

// ==================== 設定 ====================

const CONFIG = {
  // 検索クエリ（フォーム営業への返信のみを検出）
  // 件名に「AI活用レポート」を含む返信メールに限定
  GMAIL_QUERY: 'subject:"Re:" subject:"AI活用レポート" is:unread -label:処理済み',
  MAX_THREADS: 20,
  MAX_BODY_LENGTH: 800, // APIコスト管理のため冒頭800文字のみ送信
  PROCESSED_LABEL: "処理済み",
  // CRM列マッピング（0-indexed）
  COL: {
    DATE: 0,        // A: 作成日
    COMPANY: 1,     // B: 会社名
    SITE_URL: 2,    // C: ホームページURL
    ADDRESS: 3,     // D: 住所
    PHONE: 4,       // E: 電話番号
    REPORT_URL: 5,  // F: レポートURL
    OUTREACH: 6,    // G: フォーム営業文
    STATUS: 7,      // H: ステータス
    SCORE: 8,       // I: スコア
    RANK: 9,        // J: ランク
    SCORING_DATE: 10, // K: スコアリング日
    CONTACT_PATH: 11, // L: 接触経路
    RESPONSE_NOTES: 12, // M: 反応メモ
    ACTION: 13,     // N: 推奨アクション
  },
};

// ==================== メイン処理 ====================

/**
 * メインエントリポイント: 返信メールを検知してCRMに反映する
 * トリガーで5分おきに実行
 */
function checkReplyEmails() {
  const threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, CONFIG.MAX_THREADS);

  if (threads.length === 0) {
    Logger.log("新しい返信メールはありません");
    return;
  }

  const sheet = getOrCreateSheet();
  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);

  for (const thread of threads) {
    try {
      processThread(thread, sheet, processedLabel);
    } catch (e) {
      Logger.log(`Error processing thread: ${e.message}`);
    }
  }
}

/**
 * 1スレッドを処理する
 */
function processThread(thread, sheet, processedLabel) {
  const messages = thread.getMessages();
  const latestMsg = messages[messages.length - 1];

  // 送信者情報を抽出
  const fromRaw = latestMsg.getFrom();
  const emailAddr = extractEmail(fromRaw);
  const senderName = fromRaw.replace(/<.+>/, "").trim();
  const subject = latestMsg.getSubject();
  const body = latestMsg.getPlainBody().substring(0, CONFIG.MAX_BODY_LENGTH);
  const receivedDate = latestMsg.getDate();

  // 件名から元の送付先企業を特定（「Re: 株式会社〇〇様向けに...」のパターン）
  const companyFromSubject = extractCompanyFromSubject(subject);

  // CRM内で該当企業を検索
  const matchRow = findCompanyRow(sheet, companyFromSubject);

  // Claude APIで返信内容を解析
  const analysis = analyzeReplyWithAI(senderName, emailAddr, body, subject);

  if (matchRow > 0) {
    // 既存行のI〜N列を更新
    updateExistingRow(sheet, matchRow, analysis, emailAddr, receivedDate);
    Logger.log(`✅ 既存リード更新: ${companyFromSubject}（行${matchRow}）→ ランク${analysis.rank}`);
  } else {
    // 新規行として追加
    appendNewRow(sheet, analysis, emailAddr, senderName, receivedDate, body);
    Logger.log(`✅ 新規リード追加: ${analysis.company} → ランク${analysis.rank}`);
  }

  // 処理済みラベルを付与
  thread.addLabel(processedLabel);
  thread.markRead();
}

// ==================== AI解析 ====================

/**
 * Claude APIで返信メールを解析し、構造化データを返す
 */
function analyzeReplyWithAI(senderName, email, body, subject) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    throw new Error("ANTHROPIC_API_KEY がスクリプトプロパティに設定されていません");
  }

  const prompt = `以下はフォーム営業に対する返信メールです。内容を解析してJSONのみで返答してください。

送信者表示名: ${senderName}
メールアドレス: ${email}
件名: ${subject}
本文:
${body}

以下のJSON形式で返答してください（JSON以外のテキストは不要）:
{
  "company": "会社名（署名や本文から推定。わからなければ空欄）",
  "person": "担当者名（署名から抽出。わからなければ送信者表示名を使用）",
  "department": "部署名（わからなければ空欄）",
  "temperature": "前のめり or 普通 or 薄い",
  "challenge": "困りごと・課題（1〜2文で。明確でなければ空欄）",
  "score": 0から100の数値（スコアリング基準: 反応温度[前のめり+40/普通+20/薄い+5] + 接触経路メール返信+30 + メールアドレス取得+10 + 具体的課題記載+20〜40 + 「詳しく聞きたい」等+30），
  "rank": "A（70点以上） or B（40-69点） or C（39点以下）",
  "next_action": "推奨する次のアクション（1文で具体的に）"
}`;

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    payload: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 500,
      messages: [{ role: "user", content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error(`Claude API Error: ${result.error.message}`);
  }

  const text = result.content[0].text.trim();

  // JSON部分を抽出（```json ... ``` で囲まれている場合にも対応）
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error(`AI応答からJSONを抽出できませんでした: ${text.substring(0, 200)}`);
  }

  return JSON.parse(jsonMatch[0]);
}

// ==================== CRM操作 ====================

/**
 * CRMシートを取得（タブ名はスクリプトプロパティまたはデフォルト「リスト」）
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = PropertiesService.getScriptProperties().getProperty("SHEET_TAB_NAME") || "リスト";
  return ss.getSheetByName(tabName) || ss.getSheets()[0];
}

/**
 * 件名から企業名を抽出（「Re: 株式会社SEEKAD様向けに、AI活用レポート...」→ 「株式会社SEEKAD」）
 */
function extractCompanyFromSubject(subject) {
  // "Re: " を除去
  const cleaned = subject.replace(/^Re:\s*/i, "").trim();
  // "〇〇様向けに" パターン
  const match = cleaned.match(/^(.+?)様/);
  return match ? match[1] : "";
}

/**
 * CRM内でB列（会社名）を検索し、一致する行番号（1-indexed）を返す
 */
function findCompanyRow(sheet, companyName) {
  if (!companyName) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const companyCol = sheet.getRange(2, CONFIG.COL.COMPANY + 1, lastRow - 1, 1).getValues();
  let lastMatch = -1;
  for (let i = 0; i < companyCol.length; i++) {
    const cell = String(companyCol[i][0]);
    if (cell.includes(companyName) || companyName.includes(cell)) {
      lastMatch = i + 2; // 1-indexed, ヘッダー分+1（最後にマッチした行を使う）
    }
  }
  return lastMatch;
}

/**
 * 既存行のスコアリング列（I〜N）を更新
 */
function updateExistingRow(sheet, rowNum, analysis, emailAddr, receivedDate) {
  const col = CONFIG.COL;
  const dateStr = formatDate(receivedDate);

  // I列: スコア
  sheet.getRange(rowNum, col.SCORE + 1).setValue(analysis.score);
  // J列: ランク
  sheet.getRange(rowNum, col.RANK + 1).setValue(analysis.rank);
  // K列: スコアリング日
  sheet.getRange(rowNum, col.SCORING_DATE + 1).setValue(dateStr);
  // L列: 接触経路
  sheet.getRange(rowNum, col.CONTACT_PATH + 1).setValue("メール返信");
  // M列: 反応メモ
  const memo = [
    analysis.temperature ? `温度: ${analysis.temperature}` : "",
    analysis.person ? `担当: ${analysis.person}` : "",
    analysis.department ? `部署: ${analysis.department}` : "",
    analysis.challenge || "",
  ].filter(Boolean).join(" / ");
  sheet.getRange(rowNum, col.RESPONSE_NOTES + 1).setValue(memo);
  // N列: 推奨アクション
  sheet.getRange(rowNum, col.ACTION + 1).setValue(analysis.next_action || "");

  // H列: ステータス更新（ランクに応じて）
  const statusMap = { A: "Aランク対応中", B: "ナーチャリング中", C: "3ヶ月後フォロー" };
  if (statusMap[analysis.rank]) {
    sheet.getRange(rowNum, col.STATUS + 1).setValue(statusMap[analysis.rank]);
  }
}

/**
 * 新規行としてCRMに追加
 */
function appendNewRow(sheet, analysis, emailAddr, senderName, receivedDate, body) {
  const col = CONFIG.COL;
  const dateStr = formatDate(receivedDate);
  const lastRow = sheet.getLastRow() + 1;

  // A〜N列に一括書き込み
  const rowData = [
    dateStr,                                    // A: 作成日
    analysis.company || senderName,             // B: 会社名
    "",                                         // C: ホームページURL（不明）
    "",                                         // D: 住所
    "",                                         // E: 電話番号
    "",                                         // F: レポートURL
    "",                                         // G: フォーム営業文
    analysis.rank === "A" ? "Aランク対応中" :
      analysis.rank === "B" ? "ナーチャリング中" : "3ヶ月後フォロー", // H: ステータス
    analysis.score,                             // I: スコア
    analysis.rank,                              // J: ランク
    dateStr,                                    // K: スコアリング日
    "メール返信",                                // L: 接触経路
    [
      analysis.temperature ? `温度: ${analysis.temperature}` : "",
      analysis.person ? `担当: ${analysis.person}` : "",
      analysis.challenge || "",
      `返信抜粋: ${body.substring(0, 100)}`,
    ].filter(Boolean).join(" / "),              // M: 反応メモ
    analysis.next_action || "",                 // N: 推奨アクション
  ];

  sheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);
}

// ==================== ユーティリティ ====================

/**
 * メールアドレスを抽出（"山田太郎 <yamada@example.com>" → "yamada@example.com"）
 */
function extractEmail(fromStr) {
  const match = fromStr.match(/<(.+?)>/);
  return match ? match[1] : fromStr.trim();
}

/**
 * 日付を YY_MM_DD 形式にフォーマット
 */
function formatDate(date) {
  const d = new Date(date);
  const yy = String(d.getFullYear()).slice(2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yy}_${mm}_${dd}`;
}

/**
 * Gmailラベルを取得、なければ作成
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

// ==================== テスト用 ====================

/**
 * 手動テスト: 最新の返信メール1件を処理する
 */
function testProcessLatestReply() {
  const threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, 1);
  if (threads.length === 0) {
    Logger.log("返信メールが見つかりません");
    return;
  }

  const sheet = getOrCreateSheet();
  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  processThread(threads[0], sheet, processedLabel);
  Logger.log("テスト完了");
}
