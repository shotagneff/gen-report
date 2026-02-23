/**
 * nurture-sender: ナーチャリングメール自動送信
 *
 * トリガー設定: sendNurtureEmails を毎日9:00〜10:00に実行
 *
 * 処理:
 * 1. ナーチャリングタブをスキャン
 * 2. 送信予定日が今日以前 & 未送信 & ナーチャリング停止でない → 対象
 * 3. トラッキングURL埋め込み → GmailApp.sendEmail で送信
 * 4. 送信日時を記録、アクティビティに記録
 */

// ==================== メイン処理 ====================

function sendNurtureEmails() {
  const sheet = getNurtureSheet_();
  if (!sheet) {
    Logger.log("ナーチャリングタブが見つかりません");
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("送信対象なし");
    return;
  }

  const today = formatNurtureDate_();
  let sentCount = 0;
  let skipCount = 0;

  // 停止中の企業をキャッシュ（API呼び出し削減）
  const stoppedCompanies = new Set();

  for (let i = 1; i < data.length; i++) {
    const company = String(data[i][NURTURE_CONFIG.COL.COMPANY] || "");
    const email = String(data[i][NURTURE_CONFIG.COL.EMAIL] || "");
    const step = parseInt(String(data[i][NURTURE_CONFIG.COL.STEP] || "0"), 10);
    const subject = String(data[i][NURTURE_CONFIG.COL.SUBJECT] || "");
    const body = String(data[i][NURTURE_CONFIG.COL.BODY] || "");
    const scheduled = String(data[i][NURTURE_CONFIG.COL.SCHEDULED] || "");
    const sentAt = String(data[i][NURTURE_CONFIG.COL.SENT_AT] || "");

    // 既に送信済みならスキップ
    if (sentAt) continue;

    // 送信予定日が未来ならスキップ
    if (scheduled > today) continue;

    // 会社名やメールが空ならスキップ
    if (!company || !email) continue;

    // 停止中チェック（キャッシュ済み）
    if (stoppedCompanies.has(company)) {
      skipCount++;
      continue;
    }

    // リストタブのステータスチェック
    if (isNurtureStopped_(company)) {
      stoppedCompanies.add(company);
      skipCount++;
      Logger.log("停止中のためスキップ: " + company);
      continue;
    }

    // 前のStepが未送信ならスキップ（順序保証）
    if (step > 1 && !isPreviousStepSent_(data, company, step)) {
      continue;
    }

    // トラッキングURL埋め込み
    const htmlBody = buildTrackingHtml_(body, company, step);

    // メール送信
    try {
      const senderName = NURTURE_CONFIG.SENDER_NAME;
      GmailApp.sendEmail(email, subject, body, {
        htmlBody: htmlBody,
        name: senderName || undefined,
      });

      // 送信日時を記録
      const now = formatNurtureDate_(new Date()) + " " +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
      sheet.getRange(i + 1, NURTURE_CONFIG.COL.SENT_AT + 1).setValue(now);

      // アクティビティ記録
      logNurtureActivity_(
        company,
        "メール送信",
        "ナーチャリング Step" + step + " 送信完了（件名: " + subject + "）",
        "送信先: " + email
      );

      // リストタブのV列（最終接触日）を更新
      updateLastContactDate_(company);

      sentCount++;
      Logger.log("送信完了: " + company + " Step" + step);

    } catch (err) {
      Logger.log("送信エラー: " + company + " Step" + step + " - " + err.message);
    }

    // GAS実行時間制限を考慮して少し待つ
    Utilities.sleep(1000);
  }

  Logger.log("送信完了: " + sentCount + "件, スキップ: " + skipCount + "件");
}

// ==================== トラッキングHTML構築 ====================

/**
 * メール本文にトラッキング要素を埋め込む
 */
function buildTrackingHtml_(plainBody, company, step) {
  const webappUrl = NURTURE_CONFIG.WEBAPP_URL;
  if (!webappUrl) {
    // Web AppのURLが未設定ならプレーンテキストをHTMLに変換するだけ
    return plainBody.replace(/\n/g, "<br>");
  }

  const encodedCompany = encodeURIComponent(company);

  // プレーンテキストをHTMLに変換
  let html = plainBody.replace(/\n/g, "<br>");

  // リンクをトラッキングリダイレクトに変換
  html = html.replace(
    /https?:\/\/[^\s<>"]+/g,
    function(url) {
      const trackUrl = webappUrl +
        "?type=click&company=" + encodedCompany +
        "&step=" + step +
        "&url=" + encodeURIComponent(url);
      return '<a href="' + trackUrl + '">' + url + '</a>';
    }
  );

  // 開封トラッキングピクセルを末尾に追加
  const pixelUrl = webappUrl +
    "?type=open&company=" + encodedCompany +
    "&step=" + step;
  html += '<img src="' + pixelUrl + '" width="1" height="1" style="display:none" alt="">';

  return html;
}

// ==================== ヘルパー ====================

/**
 * ナーチャリングが停止中かどうかを判定
 */
function isNurtureStopped_(companyName) {
  const info = getCompanyStatus_(companyName);
  if (!info) return false;
  return NURTURE_CONFIG.STOP_STATUSES.includes(info.status);
}

/**
 * 前のStepが送信済みかどうかを確認
 */
function isPreviousStepSent_(allData, company, currentStep) {
  for (let i = 1; i < allData.length; i++) {
    const rowCompany = String(allData[i][NURTURE_CONFIG.COL.COMPANY] || "");
    const rowStep = parseInt(String(allData[i][NURTURE_CONFIG.COL.STEP] || "0"), 10);
    const rowSent = String(allData[i][NURTURE_CONFIG.COL.SENT_AT] || "");

    if (rowStep === currentStep - 1 &&
        (rowCompany.includes(company) || company.includes(rowCompany))) {
      return !!rowSent;
    }
  }
  return false;
}

/**
 * リストタブのV列（最終接触日）を更新
 */
function updateLastContactDate_(companyName) {
  try {
    const info = getCompanyStatus_(companyName);
    if (!info) return;

    const listSheet = getListSheet_();
    if (!listSheet) return;

    listSheet.getRange(info.rowIndex, NURTURE_CONFIG.LIST_COL.LAST_CONTACT + 1)
      .setValue(formatNurtureDate_());
  } catch (_) {
    // 更新失敗は無視
  }
}

// ==================== テスト用 ====================

/**
 * 手動テスト: 送信対象の一覧を表示（実際には送信しない）
 */
function testListPendingEmails() {
  const sheet = getNurtureSheet_();
  if (!sheet) {
    Logger.log("ナーチャリングタブが見つかりません");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const today = formatNurtureDate_();

  Logger.log("=== 送信待ちメール一覧 ===");
  Logger.log("今日: " + today);

  for (let i = 1; i < data.length; i++) {
    const company = String(data[i][NURTURE_CONFIG.COL.COMPANY] || "");
    const step = String(data[i][NURTURE_CONFIG.COL.STEP] || "");
    const scheduled = String(data[i][NURTURE_CONFIG.COL.SCHEDULED] || "");
    const sentAt = String(data[i][NURTURE_CONFIG.COL.SENT_AT] || "");
    const subject = String(data[i][NURTURE_CONFIG.COL.SUBJECT] || "");

    if (!sentAt && scheduled <= today) {
      const stopped = isNurtureStopped_(company);
      Logger.log(
        (stopped ? "[停止] " : "[送信対象] ") +
        company + " Step" + step +
        " 予定: " + scheduled +
        " 件名: " + subject
      );
    }
  }
}
