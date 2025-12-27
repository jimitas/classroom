/**
 * スプレッドシート操作モジュール
 * @module SheetService
 */

/**
 * シートのデータを取得
 * @param {string} sheetName - シート名
 * @returns {Array<Array>} シートデータ（2次元配列）
 */
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }

    return sheet.getDataRange().getValues();
  } catch (error) {
    console.error(`シートデータ取得エラー: ${sheetName}`, error);
    throw error;
  }
}

/**
 * シートに行を追加
 * @param {string} sheetName - シート名
 * @param {Array} rowData - 追加する行データ
 */
function appendToSheet(sheetName, rowData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }

    sheet.appendRow(rowData);
  } catch (error) {
    console.error(`シート追加エラー: ${sheetName}`, error);
    throw error;
  }
}

/**
 * クラスマスタから指定された状態のクラスを取得
 * @param {string} classState - クラス状態（CONFIG.CLASS_STATE.*）
 * @returns {Array<Object>} クラス情報の配列
 */
function getClassesByState(classState) {
  const data = getSheetData(CONFIG.SHEET_NAMES.CLASS_MASTER);
  const classes = [];

  // ヘッダー行をスキップ（1行目）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // 列の定義
    const subjectName = row[0];        // A: 科目名
    const classroomId = row[1];        // B: クラスID
    const enrollmentColumn = row[2];   // C: 履修登録マスタの列名
    const isActive = row[3];           // D: 科目有効性
    const state = row[4];              // E: クラス状態
    const teacherName = row[5];        // F: 担当教員名
    const ownerEmail = row[6];         // G: 最終オーナーメールアドレス
    const createTopics = row[7];       // H: トピック作成

    // 指定された状態のクラスのみ抽出
    if (state === classState) {
      classes.push({
        subjectName: subjectName,
        classroomId: classroomId,
        enrollmentColumn: enrollmentColumn,
        isActive: isActive,
        state: state,
        teacherName: teacherName,
        ownerEmail: ownerEmail,
        createTopics: createTopics,   // トピック作成フラグ
        rowIndex: i // 更新用に行番号を保持
      });
    }
  }

  debugLog(`クラス状態「${classState}」のクラス数: ${classes.length}`);
  return classes;
}

/**
 * クラスマスタの特定行を更新
 * @param {number} rowIndex - 行インデックス（0始まり）
 * @param {Object} updates - 更新内容
 * @param {string} updates.classroomId - クラスID（オプション）
 * @param {string} updates.state - クラス状態（オプション）
 */
function updateClassMaster(rowIndex, updates) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CLASS_MASTER);

    if (!sheet) {
      throw new Error(`シート「${CONFIG.SHEET_NAMES.CLASS_MASTER}」が見つかりません`);
    }

    // 実際の行番号は+1（スプレッドシートは1始まり）
    const actualRow = rowIndex + 1;

    // クラスIDの更新（B列）
    if (updates.classroomId !== undefined) {
      sheet.getRange(actualRow, 2).setValue(updates.classroomId);
    }

    // クラス状態の更新（E列）
    if (updates.state !== undefined) {
      sheet.getRange(actualRow, 5).setValue(updates.state);
    }

    debugLog(`クラスマスタ更新: 行${actualRow}`, updates);
  } catch (error) {
    console.error("クラスマスタ更新エラー:", error);
    throw error;
  }
}

/**
 * 履修登録マスタから生徒の履修情報を取得（縦持ち変換）
 * @returns {Array<Object>} 履修情報の配列 [{studentId, studentName, subjectName}, ...]
 */
function getEnrollmentData() {
  const data = getSheetData(CONFIG.SHEET_NAMES.ENROLLMENT_MASTER);
  const enrollments = [];

  if (data.length < 2) {
    return enrollments; // データなし
  }

  const headers = data[0]; // 1行目: ヘッダー（氏名, 学籍番号, 科目名...）

  // データ行をループ（2行目から）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const studentName = row[0];   // A: 氏名
    const studentId = row[1];     // B: 学籍番号

    // 科目列をループ（C列から）
    for (let j = 2; j < row.length; j++) {
      if (row[j] === "○") {
        enrollments.push({
          studentName: studentName,
          studentId: studentId,
          subjectName: headers[j]
        });
      }
    }
  }

  debugLog(`履修登録データ件数: ${enrollments.length}`);
  return enrollments;
}

/**
 * 教員マスタから教員の担当情報を取得（縦持ち変換）
 * @returns {Array<Object>} 担当情報の配列 [{teacherName, teacherEmail, isActive, subjectName}, ...]
 */
function getTeacherData() {
  const data = getSheetData(CONFIG.SHEET_NAMES.TEACHER_MASTER);
  const teachers = [];

  if (data.length < 2) {
    return teachers; // データなし
  }

  const headers = data[0]; // 1行目: ヘッダー（教員氏名, 教員メールアドレス, 教員アカウント有効性, 科目名...）

  // データ行をループ（2行目から）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const teacherName = row[0];    // A: 教員氏名
    const teacherEmail = row[1];   // B: 教員メールアドレス
    const isActive = row[2];       // C: 教員アカウント有効性

    // 科目列をループ（D列から）
    for (let j = 3; j < row.length; j++) {
      if (row[j] === "○") {
        teachers.push({
          teacherName: teacherName,
          teacherEmail: teacherEmail,
          isActive: isActive,
          subjectName: headers[j]
        });
      }
    }
  }

  debugLog(`教員担当データ件数: ${teachers.length}`);
  return teachers;
}

/**
 * アカウントマッピングから学籍番号に対応するGoogleメールアドレスを取得
 * @param {string} studentId - 学籍番号
 * @returns {Object|null} {name, email, userId, isActive} または null
 */
function getStudentAccount(studentId) {
  const data = getSheetData(CONFIG.SHEET_NAMES.ACCOUNT_MAPPING);

  // ヘッダー行をスキップ（1行目）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (row[0] === studentId) { // A: 内部ID（学籍番号）
      return {
        name: row[1],        // B: 氏名
        email: row[2],       // C: Googleメールアドレス
        userId: row[3],      // D: GoogleユーザーID
        isActive: row[5]     // F: アカウント有効性
      };
    }
  }

  return null;
}

/**
 * 履修登録マスタを取得
 * @returns {Array<Array>} 履修登録マスタのデータ（2次元配列）
 */
function getEnrollmentMaster() {
  return getSheetData(CONFIG.SHEET_NAMES.ENROLLMENT_MASTER);
}

/**
 * 教員マスタを取得
 * @returns {Array<Array>} 教員マスタのデータ（2次元配列）
 */
function getTeacherMaster() {
  return getSheetData(CONFIG.SHEET_NAMES.TEACHER_MASTER);
}

/**
 * アカウントマッピングを取得してマップ化
 * @returns {Object} 学籍番号 → アカウント情報のマップ
 */
function getAccountMappingMap() {
  const data = getSheetData(CONFIG.SHEET_NAMES.ACCOUNT_MAPPING);
  const accountMap = {};

  // ヘッダー行をスキップ（1行目）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const studentId = row[0];    // A: 内部ID（学籍番号）
    const name = row[1];          // B: 氏名
    const email = row[2];         // C: Googleメールアドレス
    const userId = row[3];        // D: GoogleユーザーID
    const isActive = row[5];      // F: アカウント有効性

    if (studentId) {
      accountMap[studentId] = {
        name: name,
        email: email,
        userId: userId,
        isActive: isActive
      };
    }
  }

  return accountMap;
}
