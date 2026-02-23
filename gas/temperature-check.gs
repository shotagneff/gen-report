/**
 * temperature-check: 温度上昇検知 & アラート発動
 *
 * トリガー設定: checkTemperature を1時間おきに実行
 *
 * 処理:
 * 1. ナーチャリングタブの開封・クリックデータを集計
 * 2. 会社ごとに温度スコアを計算
 * 3. 60点以上 → アラートメール送信 + CRM更新 + アクティビティ記録
 */

// ==================== メイン処理 ====================

function checkTemperature() {
  const sheet = getNurtureSheet_();
  if (!sheet) {
    Logger.log("ナーチャリングタブが見つかりません");
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  // 会社ごとにデータをグルーピング
  const companyMap = {};
  for (let i = 1; i < data.length; i++) {
    const company = String(data[i][NURTURE_CONFIG.COL.COMPANY] || "");
    if (!company) continue;

    if (!companyMap[company]) {
      companyMap[company] = [];
    }

    companyMap[company].push({
      step: parseInt(String(data[i][NURTURE_CONFIG.COL.STEP] || "0"), 10),
      sentAt: String(data[i][NURTURE_CONFIG.COL.SENT_AT] || ""),
      opened: String(data[i][NURTURE_CONFIG.COL.OPENED] || "").toUpperCase() === "TRUE",
      clicks: parseInt(String(data[i][NURTURE_CONFIG.COL.CLICKS] || "0"), 10),
    });
  }

  // 各企業の温度スコアを計算
  let alertCount = 0;

  for (const company in companyMap) {
    // 既にアラート済み or 停止中ならスキップ
    const companyInfo = getCompanyStatus_(company);
    if (companyInfo && NURTURE_CONFIG.STOP_STATUSES.includes(companyInfo.status)) {
      continue;
    }

    const result = calculateTemperatureScore_(companyMap[company], company);

    if (result.score >= NURTURE_CONFIG.TEMPERATURE.THRESHOLD) {
      triggerTemperatureAlert_(company, result.score, result.details);
      alertCount++;
    }
  }

  Logger.log("温度チェック完了: " + Object.keys(companyMap).length + "社チェック、" + alertCount + "件アラート");
}

// ==================== スコア計算 ====================

/**
 * 温度スコアを計算する
 * @param {Array} rows - 該当企業の全ナーチャリング行データ
 * @param {string} companyName - 会社名
 * @returns {{ score: number, details: string }}
 */
function calculateTemperatureScore_(rows, companyName) {
  var score = 0;
  var details = [];

  // (1) Step1〜3をすべて開封 → +30
  var steps123 = rows.filter(function(r) { return r.step >= 1 && r.step <= 3 && r.sentAt; });
  if (steps123.length === 3) {
    var allOpened = steps123.every(function(r) { return r.opened; });
    if (allOpened) {
      score += NURTURE_CONFIG.TEMPERATURE.ALL_STEPS_OPENED;
      details.push("Step1~3全開封: +" + NURTURE_CONFIG.TEMPERATURE.ALL_STEPS_OPENED);
    }
  }

  // (2) 全Stepのクリック合計
  var totalClicks = rows.reduce(function(sum, r) { return sum + r.clicks; }, 0);
  if (totalClicks >= 2) {
    score += NURTURE_CONFIG.TEMPERATURE.MULTI_CLICK;
    details.push(totalClicks + "回クリック: +" + NURTURE_CONFIG.TEMPERATURE.MULTI_CLICK);
  } else if (totalClicks === 1) {
    score += NURTURE_CONFIG.TEMPERATURE.SINGLE_CLICK;
    details.push("1回クリック: +" + NURTURE_CONFIG.TEMPERATURE.SINGLE_CLICK);
  }

  // (3) 返信チェック（リストタブの接触経路がメール返信かどうか）
  if (hasReplyDuringNurture_(companyName)) {
    score += NURTURE_CONFIG.TEMPERATURE.REPLY;
    details.push("メール返信あり: +" + NURTURE_CONFIG.TEMPERATURE.REPLY);
  }

  return { score: score, details: details.join("\n") };
}

/**
 * ナーチャリング期間中にメール返信があったかを判定
 * アクティビティタブで「メール返信受信」を検索
 */
function hasReplyDuringNurture_(companyName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSheet = ss.getSheetByName(NURTURE_CONFIG.ACTIVITY_TAB_NAME);
  if (!actSheet) return false;

  var data = actSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var actCompany = String(data[i][1] || ""); // B列: 会社名
    var actType = String(data[i][2] || "");    // C列: 種別

    if (actType === "メール返信受信" &&
        (actCompany.includes(companyName) || companyName.includes(actCompany))) {
      return true;
    }
  }
  return false;
}

// ==================== アラート発動 ====================

/**
 * 温度上昇アラートを発動する
 */
function triggerTemperatureAlert_(companyName, score, details) {
  Logger.log("🔥 温度上昇検知: " + companyName + " (スコア: " + score + "点)");

  // 1. リストタブのステータスを更新
  updateListStatus_(companyName, "温度上昇アラート");

  // 2. メール通知
  var alertEmail = NURTURE_CONFIG.SENDER_EMAIL;
  if (alertEmail) {
    var subject = "【温度上昇】" + companyName + "（スコア: " + score + "点）";
    var body = "企業: " + companyName + "\n" +
               "温度スコア: " + score + "点（閾値: " + NURTURE_CONFIG.TEMPERATURE.THRESHOLD + "点）\n\n" +
               "【スコア内訳】\n" + details + "\n\n" +
               "即座に個別アプローチを開始してください。\n" +
               "推奨: 電話でアポ打診\n\n" +
               "「先日お送りした事例、ご覧いただけましたか？\n" +
               "あの中で御社に近いケースはありましたか？\n" +
               "もし気になるところがあれば、30分だけ話しませんか？」";

    try {
      GmailApp.sendEmail(alertEmail, subject, body);
      Logger.log("アラートメール送信: " + alertEmail);
    } catch (err) {
      Logger.log("アラートメール送信エラー: " + err.message);
    }
  }

  // 3. アクティビティ記録
  logNurtureActivity_(
    companyName,
    "スコアリング更新",
    "温度上昇検知: " + score + "点 → ナーチャリング停止・個別アプローチ開始\n" + details,
    "温度上昇アラート発動"
  );
}

// ==================== テスト用 ====================

/**
 * 手動テスト: 全企業の温度スコアを表示（アラートは発動しない）
 */
function testCheckTemperatureScores() {
  var sheet = getNurtureSheet_();
  if (!sheet) {
    Logger.log("ナーチャリングタブが見つかりません");
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("データなし");
    return;
  }

  var companyMap = {};
  for (var i = 1; i < data.length; i++) {
    var company = String(data[i][NURTURE_CONFIG.COL.COMPANY] || "");
    if (!company) continue;
    if (!companyMap[company]) companyMap[company] = [];
    companyMap[company].push({
      step: parseInt(String(data[i][NURTURE_CONFIG.COL.STEP] || "0"), 10),
      sentAt: String(data[i][NURTURE_CONFIG.COL.SENT_AT] || ""),
      opened: String(data[i][NURTURE_CONFIG.COL.OPENED] || "").toUpperCase() === "TRUE",
      clicks: parseInt(String(data[i][NURTURE_CONFIG.COL.CLICKS] || "0"), 10),
    });
  }

  Logger.log("=== 温度スコア一覧 ===");
  for (var company in companyMap) {
    var result = calculateTemperatureScore_(companyMap[company], company);
    var status = result.score >= NURTURE_CONFIG.TEMPERATURE.THRESHOLD ? "🔥 アラート対象" : "正常";
    Logger.log(company + ": " + result.score + "点 [" + status + "]");
    if (result.details) {
      Logger.log("  " + result.details.replace(/\n/g, "\n  "));
    }
  }
}
