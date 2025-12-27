/**
 * システム設定モジュール
 * @module Config
 */

/**
 * システム設定定数
 */
const CONFIG = {
  // スプレッドシートID
  SPREADSHEET_ID: "1BJ-KBSQyJqbb9F8ixBWfXKLlPP4iALfsfuv9DJ-2or4",

  // シート名定義
  SHEET_NAMES: {
    CLASS_MASTER: "クラス（科目）マスタ",
    ENROLLMENT_MASTER: "履修登録マスタ",
    TEACHER_MASTER: "教員マスタ",
    ACCOUNT_MAPPING: "アカウントマッピング",
    PROCESS_LOG: "登録処理ログ",
    ERROR_LIST: "エラー未処理リスト",
    SYSTEM_SETTINGS: "システム設定 ", // 末尾にスペースあり
    TEMP_WORKSPACE: "一時作業スペース"
  },

  // クラス状態の定義
  CLASS_STATE: {
    NEW: "New",           // 新規作成待ち
    ACTIVE: "Active",     // 運用中
    ARCHIVED: "Archived"  // アーカイブ対象
  },

  // トピック構造（逆順で定義 - この順で作成される）
  TOPIC_STRUCTURE: [
    "その他",
    "テスト",
    "課題",
    "授業資料",
    "お知らせ"
  ],

  // エラーコード
  ERROR_CODES: {
    ERR_001: "APIトークン期限切れ",
    ERR_002: "権限不足",
    ERR_003: "必須フィールド欠損",
    ERR_004: "不正なデータ形式",
    ERR_005: "メールアドレス形式不正",
    ERR_006: "Rate Limit超過",
    ERR_007: "クラス登録上限超過",
    ERR_120: "Google Workspaceアカウントなし",
    WARN_101: "招待リンク期限切れ",
    WARN_200: "履修マスタに対象科目なし"
  },

  // API制限対策
  API_SETTINGS: {
    RETRY_MAX: 3,              // 最大リトライ回数
    RETRY_WAIT_MS: 1000,       // 初回リトライ待機時間（ミリ秒）
    REQUEST_INTERVAL_MS: 300   // リクエスト間隔（ミリ秒）
  }
};

/**
 * システム設定シートから設定値を取得
 * @param {string} settingName - 設定項目名
 * @returns {string|null} 設定値
 */
function getSystemSetting(settingName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SYSTEM_SETTINGS);

    if (!sheet) {
      throw new Error(`シート「${CONFIG.SHEET_NAMES.SYSTEM_SETTINGS}」が見つかりません`);
    }

    const data = sheet.getDataRange().getValues();

    // ヘッダー行をスキップして検索（1行目はヘッダー）
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === settingName) {
        return data[i][1]; // B列（設定値）を返す
      }
    }

    return null;
  } catch (error) {
    console.error(`設定値取得エラー: ${settingName}`, error);
    throw error;
  }
}

/**
 * DRY_RUN_MODEかどうかを確認
 * @returns {boolean} DRY_RUN_MODEの場合true
 */
function isDryRunMode() {
  const value = getSystemSetting("DRY_RUN_MODE");
  return value === "TRUE" || value === true;
}

/**
 * デバッグモードかどうかを確認
 * @returns {boolean} デバッグモードの場合true
 */
function isDebugMode() {
  const value = getSystemSetting("デバッグモード");
  return value === "TRUE" || value === true;
}

/**
 * 管理者アカウントIDを取得
 * @returns {string} 管理者アカウントID
 */
function getAdminAccountId() {
  const value = getSystemSetting("管理者アカウントID");
  if (!value) {
    throw new Error("管理者アカウントIDが設定されていません");
  }
  return value;
}
