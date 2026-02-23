/**
 * nurture-config: ナーチャリング関連GASの共有設定
 *
 * reply-intake.gs と同一GASプロジェクトに配置する。
 * 全てのナーチャリングGAS（sender, tracking, temperature-check）がこの定数を参照する。
 */

// ==================== ナーチャリング設定 ====================

const NURTURE_CONFIG = {
  TAB_NAME: "ナーチャリング",
  LIST_TAB_NAME: "リスト",
  ACTIVITY_TAB_NAME: "アクティビティ",

  // ナーチャリングタブの列マッピング（0-indexed）
  COL: {
    COMPANY: 0,     // A: 会社名
    EMAIL: 1,       // B: 担当者メール
    STEP: 2,        // C: Step番号
    SUBJECT: 3,     // D: 件名
    BODY: 4,        // E: 本文
    SCHEDULED: 5,   // F: 送信予定日
    SENT_AT: 6,     // G: 送信日時
    OPENED: 7,      // H: 開封
    CLICKS: 8,      // I: クリック回数
  },
  TOTAL_COLS: 9,

  // リストタブの列マッピング（reply-intake.gs の CONFIG.COL と同じ）
  LIST_COL: {
    STATUS: 7,        // H: ステータス
    LAST_CONTACT: 21, // V: 最終接触日
  },

  // 温度スコア基準（設計書準拠、半減後）
  TEMPERATURE: {
    ALL_STEPS_OPENED: 30,    // Step1〜3をすべて開封
    SINGLE_CLICK: 25,        // リンクを1回クリック
    MULTI_CLICK: 40,         // リンクを2回以上クリック
    REPLY: 50,               // メール返信あり
    THRESHOLD: 60,           // ナーチャリング停止・個別アプローチ閾値
  },

  // ナーチャリング停止とみなすステータス
  STOP_STATUSES: [
    "温度上昇アラート",
    "停止（返信受信）",
    "Aランク対応中",
    "受注",
  ],

  // Web App URL（デプロイ後にスクリプトプロパティで設定）
  get WEBAPP_URL() {
    return PropertiesService.getScriptProperties().getProperty("NURTURE_WEBAPP_URL") || "";
  },

  // 送信者情報
  get SENDER_NAME() {
    return PropertiesService.getScriptProperties().getProperty("SENDER_NAME") || "";
  },
  get SENDER_EMAIL() {
    return PropertiesService.getScriptProperties().getProperty("SENDER_EMAIL") || "";
  },
};

// ==================== 共通ユーティリティ ====================

/**
 * ナーチャリングタブを取得する
 */
function getNurtureSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(NURTURE_CONFIG.TAB_NAME);
}

/**
 * リストタブを取得する
 */
function getListSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(NURTURE_CONFIG.LIST_TAB_NAME);
}

/**
 * アクティビティタブにログを記録する
 */
function logNurtureActivity_(companyName, activityType, content, result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actSheet = ss.getSheetByName(NURTURE_CONFIG.ACTIVITY_TAB_NAME);
  if (!actSheet) return;

  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yy_MM_dd HH:mm");

  actSheet.appendRow([
    timestamp,
    companyName,
    activityType,
    "",           // 担当者名
    content,
    result || "",
    "自動",
  ]);
}

/**
 * 日付を YY_MM_DD 形式にフォーマット
 */
function formatNurtureDate_(date) {
  const d = date || new Date();
  const yy = String(d.getFullYear()).slice(2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return yy + "_" + mm + "_" + dd;
}

/**
 * リストタブで会社名に一致する行のステータスを取得する
 */
function getCompanyStatus_(companyName) {
  const listSheet = getListSheet_();
  if (!listSheet) return null;

  const data = listSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const cell = String(data[i][1] || ""); // B列: 会社名
    if (cell.includes(companyName) || companyName.includes(cell)) {
      return {
        rowIndex: i + 1,
        status: String(data[i][NURTURE_CONFIG.LIST_COL.STATUS] || ""),
      };
    }
  }
  return null;
}

/**
 * リストタブの指定行のステータスとV列（最終接触日）を更新する
 */
function updateListStatus_(companyName, newStatus) {
  const info = getCompanyStatus_(companyName);
  if (!info) return;

  const listSheet = getListSheet_();
  if (!listSheet) return;

  listSheet.getRange(info.rowIndex, NURTURE_CONFIG.LIST_COL.STATUS + 1).setValue(newStatus);
  listSheet.getRange(info.rowIndex, NURTURE_CONFIG.LIST_COL.LAST_CONTACT + 1).setValue(formatNurtureDate_());
}
