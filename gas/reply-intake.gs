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
  GMAIL_QUERY: 'subject:"Re:" subject:"AI活用レポート" is:unread -label:処理済み',
  MAX_THREADS: 20,
  MAX_BODY_LENGTH: 800,
  PROCESSED_LABEL: "処理済み",
  // CRM列マッピング（0-indexed） 22列
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
    PIPELINE: 14,   // O: パイプライン
    DEAL_AMOUNT: 15, // P: ディール金額
    WIN_PROB: 16,   // Q: 受注確度(%)
    EXPECTED_CLOSE: 17, // R: 予想受注日
    CONTACT_NAME: 18, // S: 担当者名
    CONTACT_EMAIL: 19, // T: 担当者メール
    CONTACT_DEPT: 20, // U: 担当者部署
    LAST_CONTACT: 21, // V: 最終接触日
  },
  TOTAL_COLS: 22,
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

  // 件名から元の送付先企業を特定
  const companyFromSubject = extractCompanyFromSubject(subject);

  // CRM内で該当企業を検索
  const matchRow = findCompanyRow(sheet, companyFromSubject);

  // Claude APIで返信内容を解析
  const analysis = analyzeReplyWithAI(senderName, emailAddr, body, subject);

  if (matchRow > 0) {
    // 既存行のスコアリング列を更新
    updateExistingRow(sheet, matchRow, analysis, emailAddr, receivedDate);
    Logger.log(`✅ 既存リード更新: ${companyFromSubject}（行${matchRow}）→ ランク${analysis.rank}`);
  } else {
    // 新規行として追加
    appendNewRow(sheet, analysis, emailAddr, senderName, receivedDate, body);
    Logger.log(`✅ 新規リード追加: ${analysis.company} → ランク${analysis.rank}`);
  }

  // アクティビティを記録
  try {
    logActivity(
      analysis.company || companyFromSubject || senderName,
      "メール返信受信",
      analysis.person || senderName,
      `件名: ${subject} / 温度: ${analysis.temperature}`,
      analysis.challenge || "",
      "自動"
    );
  } catch (e) {
    Logger.log(`アクティビティ記録失敗: ${e.message}`);
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
  "score": 0から100の数値,
  "rank": "A（70点以上） or B（40-69点） or C（39点以下）",
  "next_action": "推奨する次のアクション（1文で具体的に）"
}

スコアリング基準（メール返信経路）:
【基本スコア】
- 反応温度: 前のめり+40 / 普通+20 / 薄い+5
- メールアドレス取得済み（返信なので必ず）: +20
- 困りごとを具体的に聞けた（業務名・時間まで把握）: +15

【メール返信シグナル】
- 「詳しく聞きたい」「一度会えますか」: +50〜60
- 具体的な課題を書いてきた: +40
- 「資料が欲しい」: +30
- 「タイミングが来たら」（保留だが拒絶ではない）: +15
- 部署名・役職が署名にある: +10

ランク判定: A(70点以上) / B(40〜69点) / C(39点以下)`;

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

  // JSON部分を抽出
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error(`AI応答からJSONを抽出できませんでした: ${text.substring(0, 200)}`);
  }

  return JSON.parse(jsonMatch[0]);
}

// ==================== CRM操作 ====================

/**
 * CRMシートを取得
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = PropertiesService.getScriptProperties().getProperty("SHEET_TAB_NAME") || "リスト";
  return ss.getSheetByName(tabName) || ss.getSheets()[0];
}

/**
 * 件名から企業名を抽出
 */
function extractCompanyFromSubject(subject) {
  const cleaned = subject.replace(/^Re:\s*/i, "").trim();
  const match = cleaned.match(/^(.+?)様/);
  return match ? match[1] : "";
}

/**
 * CRM内でB列（会社名）を検索し、最後にマッチした行番号（1-indexed）を返す
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
      lastMatch = i + 2;
    }
  }
  return lastMatch;
}

/**
 * 既存行のスコアリング列（I〜N）+ 担当者情報（S〜U）+ 最終接触日（V）を更新
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

  // O列: パイプライン（ランクに応じて自動設定）
  const pipelineMap = { A: "商談", B: "アプローチ中", C: "リード" };
  if (pipelineMap[analysis.rank]) {
    sheet.getRange(rowNum, col.PIPELINE + 1).setValue(pipelineMap[analysis.rank]);
  }

  // S列: 担当者名
  if (analysis.person) {
    sheet.getRange(rowNum, col.CONTACT_NAME + 1).setValue(analysis.person);
  }
  // T列: 担当者メール
  if (emailAddr) {
    sheet.getRange(rowNum, col.CONTACT_EMAIL + 1).setValue(emailAddr);
  }
  // U列: 担当者部署
  if (analysis.department) {
    sheet.getRange(rowNum, col.CONTACT_DEPT + 1).setValue(analysis.department);
  }
  // V列: 最終接触日
  sheet.getRange(rowNum, col.LAST_CONTACT + 1).setValue(dateStr);

  // H列: ステータス更新
  const statusMap = { A: "Aランク対応中", B: "ナーチャリング中", C: "3ヶ月後フォロー" };
  if (statusMap[analysis.rank]) {
    sheet.getRange(rowNum, col.STATUS + 1).setValue(statusMap[analysis.rank]);
  }
}

/**
 * 新規行としてCRMに追加（22列）
 */
function appendNewRow(sheet, analysis, emailAddr, senderName, receivedDate, body) {
  const dateStr = formatDate(receivedDate);
  const pipelineMap = { A: "商談", B: "アプローチ中", C: "リード" };

  const rowData = [
    dateStr,                                    // A: 作成日
    analysis.company || senderName,             // B: 会社名
    "",                                         // C: ホームページURL
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
    pipelineMap[analysis.rank] || "リード",      // O: パイプライン
    "",                                         // P: ディール金額
    "",                                         // Q: 受注確度
    "",                                         // R: 予想受注日
    analysis.person || senderName,              // S: 担当者名
    emailAddr,                                  // T: 担当者メール
    analysis.department || "",                  // U: 担当者部署
    dateStr,                                    // V: 最終接触日
  ];

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);
}

// ==================== アクティビティ記録 ====================

/**
 * アクティビティタブに1行追加する
 */
function logActivity(companyName, activityType, contactName, content, result, recorder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activitySheet = ss.getSheetByName("アクティビティ");
  if (!activitySheet) return; // タブが存在しない場合はスキップ

  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yy_MM_dd HH:mm");

  activitySheet.appendRow([
    timestamp,
    companyName,
    activityType,
    contactName || "",
    content || "",
    result || "",
    recorder || "自動",
  ]);
}

// ==================== ユーティリティ ====================

/**
 * メールアドレスを抽出
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
