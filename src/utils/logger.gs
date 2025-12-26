/**
 * ログ記録モジュール
 * @module Logger
 */

/**
 * 登録処理ログシートにログを記録
 * @param {Object} logData - ログデータ
 * @param {string} logData.studentName - 生徒名（オプション）
 * @param {string} logData.studentId - 学籍番号（オプション）
 * @param {string} logData.subjectName - 科目名
 * @param {string} logData.classroomId - Classroom ID
 * @param {string} logData.result - 処理結果（成功/失敗）
 * @param {string} logData.message - APIレスポンス/エラーメッセージ
 * @param {string} logData.executor - 実行者/システム
 */
function logToProcessLog(logData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROCESS_LOG);

    if (!sheet) {
      throw new Error(`シート「${CONFIG.SHEET_NAMES.PROCESS_LOG}」が見つかりません`);
    }

    const timestamp = new Date();
    const rowData = [
      timestamp,                        // A: 処理日時
      logData.studentName || "",        // B: 生徒名
      logData.studentId || "",          // C: 学籍番号
      logData.subjectName || "",        // D: 科目名
      logData.classroomId || "",        // E: Classroom ID
      logData.result || "",             // F: 処理結果
      logData.message || "",            // G: APIレスポンス/エラーメッセージ
      logData.executor || "Auto System", // H: 実行者/システム
      logData.invitationMethod || "",   // I: 招待方法
      logData.ownerStatus || ""         // J: 権限交代状況
    ];

    sheet.appendRow(rowData);

    if (isDebugMode()) {
      console.log("ログ記録:", logData);
    }
  } catch (error) {
    console.error("ログ記録エラー:", error);
    // ログ記録エラーは処理を停止しない
  }
}

/**
 * エラー未処理リストにエラーを記録
 * @param {Object} errorData - エラーデータ
 * @param {string} errorData.studentName - 生徒名（オプション）
 * @param {string} errorData.studentId - 学籍番号（オプション）
 * @param {string} errorData.subjectName - 科目名
 * @param {string} errorData.errorCode - エラーコード
 * @param {string} errorData.errorDetail - エラー詳細
 * @param {string} errorData.status - 対応状況（デフォルト: 未対応）
 */
function logToErrorList(errorData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ERROR_LIST);

    if (!sheet) {
      throw new Error(`シート「${CONFIG.SHEET_NAMES.ERROR_LIST}」が見つかりません`);
    }

    const timestamp = new Date();
    const rowData = [
      timestamp,                      // A: 発生日時
      errorData.studentName || "",    // B: 生徒名
      errorData.studentId || "",      // C: 学籍番号
      errorData.subjectName || "",    // D: 科目名
      errorData.errorCode || "",      // E: エラーコード
      errorData.errorDetail || "",    // F: エラー詳細
      errorData.status || "未対応",   // G: 対応状況
      "",                             // H: 対応担当者
      errorData.note || "",           // I: 備考
      ""                              // J: エラー発生回数
    ];

    sheet.appendRow(rowData);

    if (isDebugMode()) {
      console.log("エラー記録:", errorData);
    }
  } catch (error) {
    console.error("エラーリスト記録エラー:", error);
  }
}

/**
 * コンソールログ出力（デバッグモード時のみ）
 * @param {string} message - ログメッセージ
 * @param {*} data - ログデータ（オプション）
 */
function debugLog(message, data) {
  if (isDebugMode()) {
    if (data !== undefined) {
      console.log(message, data);
    } else {
      console.log(message);
    }
  }
}

/**
 * 処理開始ログ
 * @param {string} phaseName - フェーズ名
 */
function logPhaseStart(phaseName) {
  console.log(`========== ${phaseName} 開始 ==========`);
  debugLog(`DRY_RUN_MODE: ${isDryRunMode()}`);
}

/**
 * 処理完了ログ
 * @param {string} phaseName - フェーズ名
 * @param {number} successCount - 成功件数
 * @param {number} errorCount - エラー件数
 */
function logPhaseComplete(phaseName, successCount, errorCount) {
  console.log(`========== ${phaseName} 完了 ==========`);
  console.log(`成功: ${successCount}件, エラー: ${errorCount}件`);
}
