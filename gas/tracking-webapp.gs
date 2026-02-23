/**
 * tracking-webapp: 開封・クリックトラッキング用 Web App
 *
 * デプロイ手順:
 * 1. GASエディタ → デプロイ → ウェブアプリとしてデプロイ
 * 2. アクセスできるユーザー: 「全員」
 * 3. デプロイ後のURLをスクリプトプロパティ「NURTURE_WEBAPP_URL」に設定
 *
 * メール本文に埋め込まれるURL例:
 *   開封ピクセル: {WEBAPP_URL}?type=open&company=株式会社〇〇&step=1
 *   クリック:     {WEBAPP_URL}?type=click&company=株式会社〇〇&step=1&url=https://example.com
 */

// ==================== Web App エントリポイント ====================

function doGet(e) {
  const type = (e.parameter.type || "").toLowerCase();
  const company = decodeURIComponent(e.parameter.company || "");
  const step = parseInt(e.parameter.step || "0", 10);
  const url = decodeURIComponent(e.parameter.url || "");

  if (!company || !step) {
    return HtmlService.createHtmlOutput("OK");
  }

  try {
    if (type === "open") {
      recordOpen_(company, step);
      // 1x1 透過GIFを返す（meta refreshでData URI画像へリダイレクト）
      return HtmlService.createHtmlOutput(
        '<html><body><img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" width="1" height="1"></body></html>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (type === "click") {
      recordClick_(company, step);
      if (url) {
        // 元のURLへリダイレクト
        return HtmlService.createHtmlOutput(
          '<html><head><meta http-equiv="refresh" content="0;url=' + escapeHtml_(url) + '"></head>' +
          '<body><p>リダイレクト中... <a href="' + escapeHtml_(url) + '">こちらをクリック</a></p></body></html>'
        ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }
  } catch (err) {
    Logger.log("トラッキングエラー: " + err.message);
  }

  return HtmlService.createHtmlOutput("OK");
}

// ==================== 開封記録 ====================

function recordOpen_(companyName, step) {
  const sheet = getNurtureSheet_();
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][NURTURE_CONFIG.COL.COMPANY] || "");
    const rowStep = parseInt(String(data[i][NURTURE_CONFIG.COL.STEP] || "0"), 10);

    if (rowStep === step && (rowCompany.includes(companyName) || companyName.includes(rowCompany))) {
      // 既にTRUEの場合はスキップ（重複記録を防止）
      if (String(data[i][NURTURE_CONFIG.COL.OPENED]) === "TRUE") return;

      sheet.getRange(i + 1, NURTURE_CONFIG.COL.OPENED + 1).setValue("TRUE");

      logNurtureActivity_(
        companyName,
        "ナーチャリング開封",
        "Step" + step + " メールが開封されました",
        ""
      );
      return;
    }
  }
}

// ==================== クリック記録 ====================

function recordClick_(companyName, step) {
  const sheet = getNurtureSheet_();
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][NURTURE_CONFIG.COL.COMPANY] || "");
    const rowStep = parseInt(String(data[i][NURTURE_CONFIG.COL.STEP] || "0"), 10);

    if (rowStep === step && (rowCompany.includes(companyName) || companyName.includes(rowCompany))) {
      const currentClicks = parseInt(String(data[i][NURTURE_CONFIG.COL.CLICKS] || "0"), 10);
      sheet.getRange(i + 1, NURTURE_CONFIG.COL.CLICKS + 1).setValue(currentClicks + 1);

      logNurtureActivity_(
        companyName,
        "ナーチャリングクリック",
        "Step" + step + " リンクがクリックされました（" + (currentClicks + 1) + "回目）",
        ""
      );
      return;
    }
  }
}

// ==================== ユーティリティ ====================

function escapeHtml_(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
