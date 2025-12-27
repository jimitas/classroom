/**
 * データ変換モジュール（マトリクス→縦持ち変換）
 * @module DataTransformer
 */

/**
 * 履修登録マスタ（横持ち）から縦持ちリストを生成
 *
 * @param {Array<Array>} enrollmentData - 履修登録マスタのデータ（2次元配列）
 * @param {Object} classMasterMap - クラスマスタのマップ（科目名 → クラス情報）
 * @returns {Array<Object>} 縦持ちリスト [{studentId, studentName, subjectName, classroomId, enrollmentColumn}]
 *
 * @example
 * 入力（enrollmentData）:
 * | 氏名     | 学籍番号  | 有効 | 国語 | 数学 | 英語 |
 * |----------|----------|------|------|------|------|
 * | 田中 太郎 | S2025001 | TRUE | ○    | ○    |      |
 * | 佐藤 花子 | S2025002 | TRUE |      | ○    | ○    |
 *
 * 出力:
 * [
 *   { studentId: "S2025001", studentName: "田中 太郎", subjectName: "国語", classroomId: "...", enrollmentColumn: "国語" },
 *   { studentId: "S2025001", studentName: "田中 太郎", subjectName: "数学", classroomId: "...", enrollmentColumn: "数学" },
 *   { studentId: "S2025002", studentName: "佐藤 花子", subjectName: "数学", classroomId: "...", enrollmentColumn: "数学" },
 *   { studentId: "S2025002", studentName: "佐藤 花子", subjectName: "英語", classroomId: "...", enrollmentColumn: "英語" }
 * ]
 */
function transformEnrollmentDataToVertical(enrollmentData, classMasterMap) {
  const verticalData = [];

  if (!enrollmentData || enrollmentData.length < 2) {
    console.log("履修登録マスタにデータがありません");
    return verticalData;
  }

  // ヘッダー行（1行目）から科目名の列を取得
  const headers = enrollmentData[0];
  const studentNameCol = 0;  // A列: 氏名
  const studentIdCol = 1;    // B列: 学籍番号
  const isActiveCol = 2;     // C列: 有効フラグ
  const subjectStartCol = 3; // D列以降: 科目

  // 2行目以降がデータ行
  for (let i = 1; i < enrollmentData.length; i++) {
    const row = enrollmentData[i];
    const studentName = row[studentNameCol];
    const studentId = row[studentIdCol];
    const isActive = row[isActiveCol];

    // 学籍番号が空の場合はスキップ
    if (!studentId || studentId === "") {
      continue;
    }

    // 有効フラグがFALSEの場合はスキップ（転入生追加時の既登録者除外）
    if (isActive === false || isActive === "FALSE" || isActive === "") {
      debugLog(`スキップ: 生徒「${studentName}」(${studentId})は有効フラグ=FALSE`);
      continue;
    }

    // 各科目列をチェック
    for (let j = subjectStartCol; j < headers.length; j++) {
      const subjectColumn = headers[j];  // 列名（履修登録マスタの科目名）
      const enrollmentMark = row[j];     // セルの値（○ or 空欄）

      // "○" がある場合のみ抽出
      if (enrollmentMark === "○" || enrollmentMark === "o" || enrollmentMark === "O") {
        // クラスマスタから該当する科目を検索（C列の「履修登録マスタの列名」で照合）
        const classInfo = classMasterMap[subjectColumn];

        if (classInfo) {
          verticalData.push({
            studentId: studentId,
            studentName: studentName,
            subjectName: classInfo.subjectName,      // クラスマスタのA列（科目名）
            classroomId: classInfo.classroomId,      // クラスマスタのB列（クラスID）
            enrollmentColumn: subjectColumn,         // 履修登録マスタの列名
            isActive: classInfo.isActive             // 科目有効性
          });
        } else {
          debugLog(`警告: 履修登録マスタの列「${subjectColumn}」に対応するクラスがクラスマスタに見つかりません`);
        }
      }
    }
  }

  return verticalData;
}

/**
 * 教員マスタ（横持ち）から縦持ちリストを生成
 *
 * @param {Array<Array>} teacherData - 教員マスタのデータ（2次元配列）
 * @param {Object} classMasterMap - クラスマスタのマップ（科目名 → クラス情報）
 * @returns {Array<Object>} 縦持ちリスト [{teacherEmail, teacherName, subjectName, classroomId, isValid}]
 *
 * @example
 * 入力（teacherData）:
 * | 教員氏名  | 教員メールアドレス        | 教員アカウント有効性 | 国語 | 数学 | 英語 |
 * |----------|-------------------------|---------------------|------|------|------|
 * | 山田先生  | yamada@school.jp        | TRUE                | ○    |      | ○    |
 * | 佐藤先生  | sato@school.jp          | TRUE                |      | ○    | ○    |
 *
 * 出力:
 * [
 *   { teacherEmail: "yamada@school.jp", teacherName: "山田先生", subjectName: "国語", classroomId: "...", isValid: true },
 *   { teacherEmail: "yamada@school.jp", teacherName: "山田先生", subjectName: "英語", classroomId: "...", isValid: true },
 *   { teacherEmail: "sato@school.jp", teacherName: "佐藤先生", subjectName: "数学", classroomId: "...", isValid: true },
 *   { teacherEmail: "sato@school.jp", teacherName: "佐藤先生", subjectName: "英語", classroomId: "...", isValid: true }
 * ]
 */
function transformTeacherDataToVertical(teacherData, classMasterMap) {
  const verticalData = [];

  if (!teacherData || teacherData.length < 2) {
    console.log("教員マスタにデータがありません");
    return verticalData;
  }

  // ヘッダー行（1行目）から科目名の列を取得
  const headers = teacherData[0];
  const teacherNameCol = 0;     // A列: 教員氏名
  const teacherEmailCol = 1;    // B列: 教員メールアドレス
  const validityCol = 2;        // C列: 教員アカウント有効性
  const subjectStartCol = 3;    // D列以降: 科目

  // 2行目以降がデータ行
  for (let i = 1; i < teacherData.length; i++) {
    const row = teacherData[i];
    const teacherName = row[teacherNameCol];
    const teacherEmail = row[teacherEmailCol];
    const isValid = row[validityCol];

    // メールアドレスが空の場合はスキップ
    if (!teacherEmail || teacherEmail === "") {
      continue;
    }

    // アカウント有効性がFALSEの場合はスキップ（警告ログ）
    if (isValid === false || isValid === "FALSE") {
      debugLog(`警告: 教員「${teacherName}」(${teacherEmail})のアカウントが無効です`);
      continue;
    }

    // 各科目列をチェック
    for (let j = subjectStartCol; j < headers.length; j++) {
      const subjectName = headers[j];  // 列名（科目名）
      const assignmentMark = row[j];   // セルの値（○ or 空欄）

      // "○" がある場合のみ抽出
      if (assignmentMark === "○" || assignmentMark === "o" || assignmentMark === "O") {
        // クラスマスタから該当する科目を検索（A列の「科目名」で照合）
        const classInfo = classMasterMap[subjectName];

        if (classInfo) {
          verticalData.push({
            teacherEmail: teacherEmail,
            teacherName: teacherName,
            subjectName: subjectName,
            classroomId: classInfo.classroomId,
            isValid: isValid
          });
        } else {
          debugLog(`警告: 教員マスタの列「${subjectName}」に対応するクラスがクラスマスタに見つかりません`);
        }
      }
    }
  }

  return verticalData;
}

/**
 * クラスマスタから検索用マップを生成
 *
 * @param {Array<Object>} classes - getClassesByState()の結果
 * @returns {Object} 検索マップ（2種類）
 *   - byEnrollmentColumn: 履修登録マスタの列名 → クラス情報
 *   - bySubjectName: 科目名 → クラス情報
 */
function createClassMasterMaps(classes) {
  const byEnrollmentColumn = {};
  const bySubjectName = {};

  classes.forEach(classInfo => {
    // C列（履修登録マスタの列名）でマッピング（生徒登録用）
    if (classInfo.enrollmentColumn) {
      byEnrollmentColumn[classInfo.enrollmentColumn] = classInfo;
    }

    // A列（科目名）でマッピング（教員登録用）
    if (classInfo.subjectName) {
      bySubjectName[classInfo.subjectName] = classInfo;
    }
  });

  return {
    byEnrollmentColumn: byEnrollmentColumn,
    bySubjectName: bySubjectName
  };
}
