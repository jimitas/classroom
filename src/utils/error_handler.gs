/**
 * エラー処理モジュール
 * @module ErrorHandler
 */

/**
 * API呼び出しをリトライ付きで実行
 * @param {Function} apiFunction - 実行するAPI関数
 * @param {number} maxRetries - 最大リトライ回数（デフォルト: CONFIG.API_SETTINGS.RETRY_MAX）
 * @returns {*} API関数の戻り値
 * @throws {Error} 最大リトライ回数を超えた場合
 */
function executeWithRetry(apiFunction, maxRetries = null) {
  if (maxRetries === null) {
    maxRetries = CONFIG.API_SETTINGS.RETRY_MAX;
  }

  let retryCount = 0;
  let lastError = null;

  while (retryCount <= maxRetries) {
    try {
      return apiFunction();
    } catch (error) {
      lastError = error;
      retryCount++;

      debugLog(`エラー発生（試行 ${retryCount}/${maxRetries + 1}）`, error.message);

      // リトライ可能なエラーかチェック
      if (!isRetryableError(error)) {
        debugLog("リトライ不可能なエラーのため、処理を中断します");
        throw error;
      }

      // 最大リトライ回数に達した場合
      if (retryCount > maxRetries) {
        debugLog("最大リトライ回数に達しました");
        throw lastError;
      }

      // Rate Limitエラーの場合は指数バックオフ
      if (isRateLimitError(error)) {
        const waitTime = Math.pow(2, retryCount) * CONFIG.API_SETTINGS.RETRY_WAIT_MS;
        debugLog(`Rate Limit検出。${waitTime}ms待機します`);
        Utilities.sleep(waitTime);
      } else {
        // 通常のリトライ待機
        Utilities.sleep(CONFIG.API_SETTINGS.RETRY_WAIT_MS);
      }
    }
  }

  throw lastError;
}

/**
 * Rate Limitエラーかどうかを判定
 * @param {Error} error - エラーオブジェクト
 * @returns {boolean} Rate Limitエラーの場合true
 */
function isRateLimitError(error) {
  if (!error) return false;

  const message = error.message || error.toString();

  // Rate limitやQuota exceededなどのキーワードをチェック
  return message.includes("rate limit") ||
         message.includes("Rate Limit") ||
         message.includes("quota exceeded") ||
         message.includes("Quota exceeded") ||
         message.includes("429");
}

/**
 * リトライ可能なエラーかどうかを判定
 * @param {Error} error - エラーオブジェクト
 * @returns {boolean} リトライ可能な場合true
 */
function isRetryableError(error) {
  if (!error) return false;

  const message = error.message || error.toString();

  // リトライ可能なエラーパターン
  const retryablePatterns = [
    "rate limit",
    "Rate Limit",
    "quota exceeded",
    "Quota exceeded",
    "timeout",
    "Timeout",
    "503",
    "502",
    "429",
    "Internal error",
    "Service unavailable"
  ];

  return retryablePatterns.some(pattern => message.includes(pattern));
}

/**
 * エラーコードを特定
 * @param {Error} error - エラーオブジェクト
 * @returns {string} エラーコード
 */
function getErrorCode(error) {
  if (!error) return "UNKNOWN";

  const message = error.message || error.toString();

  // エラーメッセージからエラーコードを判定
  if (message.includes("Invalid email") || message.includes("invalid email")) {
    return "ERR_005";
  } else if (isRateLimitError(error)) {
    return "ERR_006";
  } else if (message.includes("limit exceeded") || message.includes("Maximum")) {
    return "ERR_007";
  } else if (message.includes("not found") || message.includes("does not exist")) {
    return "ERR_120";
  } else if (message.includes("permission") || message.includes("Permission")) {
    return "ERR_002";
  } else if (message.includes("token") || message.includes("Token")) {
    return "ERR_001";
  } else if (message.includes("required field") || message.includes("missing")) {
    return "ERR_003";
  } else if (message.includes("invalid") || message.includes("Invalid")) {
    return "ERR_004";
  }

  return "UNKNOWN";
}

/**
 * エラーの詳細メッセージを取得
 * @param {Error} error - エラーオブジェクト
 * @param {string} errorCode - エラーコード
 * @returns {string} エラー詳細メッセージ
 */
function getErrorDetail(error, errorCode) {
  const baseMessage = error.message || error.toString();
  const codeDescription = CONFIG.ERROR_CODES[errorCode] || "不明なエラー";

  return `${codeDescription}: ${baseMessage}`;
}

/**
 * エラーを処理してログに記録
 * @param {Error} error - エラーオブジェクト
 * @param {Object} context - エラーコンテキスト
 */
function handleError(error, context) {
  const errorCode = getErrorCode(error);
  const errorDetail = getErrorDetail(error, errorCode);

  // エラーログに記録
  logToErrorList({
    studentName: context.studentName,
    studentId: context.studentId,
    subjectName: context.subjectName,
    errorCode: errorCode,
    errorDetail: errorDetail,
    status: "未対応"
  });

  // 処理ログにも記録
  logToProcessLog({
    studentName: context.studentName,
    studentId: context.studentId,
    subjectName: context.subjectName,
    classroomId: context.classroomId,
    result: "失敗",
    message: errorDetail,
    executor: "Auto System"
  });

  console.error(`エラー処理完了: ${errorCode} - ${errorDetail}`);
}
