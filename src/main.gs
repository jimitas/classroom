/**
 * Google Classroom 年度更新自動化システム
 * メインエントリーポイント
 * @module Main
 */

/**
 * 全フェーズを順次実行
 */
function runAllPhases() {
  console.log("=".repeat(60));
  console.log("Google Classroom 年度更新自動化システム");
  console.log("全フェーズ実行開始");
  console.log("=".repeat(60));

  const results = {
    phase1: null,
    phase2: null,
    phase3: null,
    phase4: null
  };

  try {
    // Phase 1: アーカイブ
    results.phase1 = runPhase1Archive();

    // Phase 2, 3, 4は今後実装予定
    console.log("\n⚠️ Phase 2, 3, 4は未実装です");

    // 全体結果のサマリー
    console.log("\n" + "=".repeat(60));
    console.log("全フェーズ実行完了");
    console.log("=".repeat(60));
    logExecutionSummary(results);

  } catch (error) {
    console.error("全フェーズ実行中にエラーが発生しました:", error);
    throw error;
  }
}

/**
 * Phase 1のみを実行（テスト用）
 */
function testPhase1() {
  console.log("Phase 1テスト実行");
  const result = runPhase1Archive();
  console.log("Phase 1テスト完了:", result);
}

/**
 * DRY-RUNモードで全フェーズを実行（テスト用）
 */
function runDryRun() {
  console.log("⚠️ DRY-RUNモード実行");
  console.log("システム設定シートでDRY_RUN_MODE=TRUEに設定されていることを確認してください");
  console.log("");

  const dryRunMode = isDryRunMode();

  if (!dryRunMode) {
    console.log("❌ エラー: DRY_RUN_MODEがTRUEに設定されていません");
    console.log("システム設定シートでDRY_RUN_MODE=TRUEに設定してください");
    return;
  }

  console.log("✓ DRY_RUN_MODE確認完了");
  runAllPhases();
}

/**
 * システム設定のテスト（デバッグ用）
 */
function testSystemSettings() {
  console.log("システム設定テスト");
  console.log("-".repeat(40));

  try {
    console.log("スプレッドシートID:", CONFIG.SPREADSHEET_ID);
    console.log("DRY_RUN_MODE:", isDryRunMode());
    console.log("デバッグモード:", isDebugMode());

    try {
      const adminId = getAdminAccountId();
      console.log("管理者アカウントID:", adminId);
    } catch (error) {
      console.log("管理者アカウントID: 未設定");
    }

    console.log("-".repeat(40));
    console.log("✓ システム設定テスト完了");

  } catch (error) {
    console.error("システム設定テストエラー:", error);
  }
}

/**
 * スプレッドシート接続テスト（デバッグ用）
 */
function testSpreadsheetConnection() {
  console.log("スプレッドシート接続テスト");
  console.log("-".repeat(40));

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    console.log("✓ スプレッドシート接続成功");
    console.log("スプレッドシート名:", ss.getName());

    // 各シートの存在確認
    for (const [key, sheetName] of Object.entries(CONFIG.SHEET_NAMES)) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const rowCount = sheet.getLastRow();
        console.log(`✓ ${sheetName}: ${rowCount}行`);
      } else {
        console.log(`❌ ${sheetName}: 見つかりません`);
      }
    }

    console.log("-".repeat(40));
    console.log("✓ スプレッドシート接続テスト完了");

  } catch (error) {
    console.error("スプレッドシート接続テストエラー:", error);
  }
}

/**
 * クラスマスタデータのテスト表示（デバッグ用）
 */
function testClassMasterData() {
  console.log("クラスマスタデータテスト");
  console.log("-".repeat(40));

  try {
    const archivedClasses = getClassesByState(CONFIG.CLASS_STATE.ARCHIVED);
    console.log(`アーカイブ対象クラス: ${archivedClasses.length}件`);
    archivedClasses.forEach(c => {
      console.log(`  - ${c.subjectName} (${c.classroomId})`);
    });

    const newClasses = getClassesByState(CONFIG.CLASS_STATE.NEW);
    console.log(`新規作成待ちクラス: ${newClasses.length}件`);
    newClasses.forEach(c => {
      console.log(`  - ${c.subjectName}`);
    });

    const activeClasses = getClassesByState(CONFIG.CLASS_STATE.ACTIVE);
    console.log(`運用中クラス: ${activeClasses.length}件`);
    activeClasses.forEach(c => {
      console.log(`  - ${c.subjectName} (${c.classroomId})`);
    });

    console.log("-".repeat(40));
    console.log("✓ クラスマスタデータテスト完了");

  } catch (error) {
    console.error("クラスマスタデータテストエラー:", error);
  }
}

/**
 * 実行サマリーをログ出力
 * @param {Object} results - 各フェーズの実行結果
 */
function logExecutionSummary(results) {
  console.log("\n【実行サマリー】");

  if (results.phase1) {
    console.log(`Phase 1: 成功 ${results.phase1.successCount}件, エラー ${results.phase1.errorCount}件`);
  }

  if (results.phase2) {
    console.log(`Phase 2: 成功 ${results.phase2.successCount}件, エラー ${results.phase2.errorCount}件`);
  }

  if (results.phase3) {
    console.log(`Phase 3: 成功 ${results.phase3.successCount}件, エラー ${results.phase3.errorCount}件`);
  }

  if (results.phase4) {
    console.log(`Phase 4: 成功 ${results.phase4.successCount}件, エラー ${results.phase4.errorCount}件`);
  }

  const totalSuccess = (results.phase1?.successCount || 0) +
                       (results.phase2?.successCount || 0) +
                       (results.phase3?.successCount || 0) +
                       (results.phase4?.successCount || 0);

  const totalError = (results.phase1?.errorCount || 0) +
                     (results.phase2?.errorCount || 0) +
                     (results.phase3?.errorCount || 0) +
                     (results.phase4?.errorCount || 0);

  console.log(`\n合計: 成功 ${totalSuccess}件, エラー ${totalError}件`);
}

/**
 * 招待コードからクラスIDを取得（デバッグ用）
 * @param {string} enrollmentCode - 招待コード
 * @returns {string|null} クラスID
 */
function getCourseIdFromEnrollmentCode(enrollmentCode) {
  try {
    console.log(`招待コード「${enrollmentCode}」からクラスIDを検索中...`);

    // すべてのクラスを取得
    const courses = Classroom.Courses.list({
      teacherId: 'me',
      pageSize: 100
    });

    if (courses.courses && courses.courses.length > 0) {
      // 招待コードが一致するクラスを検索
      const targetCourse = courses.courses.find(course =>
        course.enrollmentCode === enrollmentCode
      );

      if (targetCourse) {
        console.log("✓ クラスが見つかりました！");
        console.log("-".repeat(60));
        console.log("クラス名:", targetCourse.name);
        console.log("クラスID:", targetCourse.id);
        console.log("招待コード:", targetCourse.enrollmentCode);
        console.log("クラス状態:", targetCourse.courseState);
        console.log("説明:", targetCourse.descriptionHeading || "なし");
        console.log("-".repeat(60));
        console.log("\n📋 クラスマスタに登録する場合:");
        console.log(`  科目名: ${targetCourse.name}`);
        console.log(`  クラスID: ${targetCourse.id}`);
        console.log(`  クラス状態: Archived （アーカイブテストの場合）`);
        return targetCourse.id;
      } else {
        console.log(`❌ 招待コード「${enrollmentCode}」に一致するクラスが見つかりません`);
        console.log("あなたが教師として参加しているクラスを確認してください。");
        console.log("→ listAllMyCourses() を実行してください");
      }
    } else {
      console.log("❌ クラスが見つかりません（あなたが教師のクラスがありません）");
    }

    return null;
  } catch (error) {
    console.error("エラー:", error);
    return null;
  }
}

/**
 * 自分が教師として参加しているすべてのクラスを表示
 */
function listAllMyCourses() {
  try {
    console.log("あなたが教師として参加しているクラスを取得中...");

    const courses = Classroom.Courses.list({
      teacherId: 'me',
      pageSize: 100
    });

    if (courses.courses && courses.courses.length > 0) {
      console.log("\n" + "=".repeat(80));
      console.log("あなたのクラス一覧");
      console.log("=".repeat(80));

      courses.courses.forEach((course, index) => {
        console.log(`\n【${index + 1}】 ${course.name}`);
        console.log(`  クラスID: ${course.id}`);
        console.log(`  招待コード: ${course.enrollmentCode || 'なし'}`);
        console.log(`  状態: ${course.courseState}`);
        console.log(`  セクション: ${course.section || 'なし'}`);
        console.log(`  説明: ${course.descriptionHeading || 'なし'}`);
        if (course.enrollmentCode) {
          console.log(`  URL: https://classroom.google.com/r/${course.enrollmentCode}`);
        }
      });

      console.log("\n" + "=".repeat(80));
      console.log(`合計: ${courses.courses.length}クラス`);
      console.log("=".repeat(80));
    } else {
      console.log("❌ クラスが見つかりません");
      console.log("あなたが教師として参加しているクラスがありません。");
    }
  } catch (error) {
    console.error("エラー:", error);
    console.error("Classroom APIへのアクセスに失敗しました。");
    console.error("権限の確認が必要な場合があります。");
  }
}

/**
 * Classroomの現在の状態をクラスマスタに反映
 * - 既存のクラスID（B列）と照合して更新
 * - 新しいクラスは追加
 * - 不明な列（C, F, G列）は既存値を保持または空白
 */
function syncClassMasterFromClassroom() {
  try {
    console.log("=".repeat(80));
    console.log("Classroomの状態をクラスマスタに反映");
    console.log("=".repeat(80));

    // Classroom APIから現在のクラスを取得
    const courses = Classroom.Courses.list({
      teacherId: 'me',
      pageSize: 100
    });

    if (!courses.courses || courses.courses.length === 0) {
      console.log("❌ クラスが見つかりません");
      return;
    }

    console.log(`取得したクラス数: ${courses.courses.length}`);

    // スプレッドシートを開く
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CLASS_MASTER);

    if (!sheet) {
      throw new Error(`シート「${CONFIG.SHEET_NAMES.CLASS_MASTER}」が見つかりません`);
    }

    // 既存データを取得
    const existingData = sheet.getDataRange().getValues();
    const headers = existingData[0];

    // 既存データをクラスIDでマップ化（B列 = クラスID）
    const existingMap = new Map();
    for (let i = 1; i < existingData.length; i++) {
      const classroomId = existingData[i][1]; // B列
      if (classroomId) {
        existingMap.set(classroomId, {
          rowIndex: i,
          data: existingData[i]
        });
      }
    }

    let updateCount = 0;
    let addCount = 0;

    // 各クラスを処理
    courses.courses.forEach(course => {
      // courseStateをマッピング（ACTIVE → Active, ARCHIVED → Archived）
      const mappedState = mapCourseState(course.courseState);

      if (existingMap.has(course.id)) {
        // 既存のクラスを更新
        const existing = existingMap.get(course.id);
        const rowIndex = existing.rowIndex;
        const existingRow = existing.data;

        // 更新する行データ（既存の不明な列は保持）
        const updatedRow = [
          course.name,                    // A: 科目名
          course.id,                      // B: クラスID
          existingRow[2] || course.name,  // C: 履修登録マスタの列名（既存値を保持、なければ科目名）
          existingRow[3] || true,         // D: 科目有効性（既存値を保持、なければTRUE）
          mappedState,                    // E: クラス状態
          existingRow[5] || "",           // F: 担当教員名（既存値を保持）
          existingRow[6] || ""            // G: 最終オーナーメールアドレス（既存値を保持）
        ];

        // スプレッドシートの行番号は1始まり
        sheet.getRange(rowIndex + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
        console.log(`✓ 更新: ${course.name} (${course.id}) - ${mappedState}`);
        updateCount++;

      } else {
        // 新しいクラスを追加
        const newRow = [
          course.name,                    // A: 科目名
          course.id,                      // B: クラスID
          course.name,                    // C: 履修登録マスタの列名（科目名と同じ）
          true,                           // D: 科目有効性
          mappedState,                    // E: クラス状態
          "",                             // F: 担当教員名（空白）
          ""                              // G: 最終オーナーメールアドレス（空白）
        ];

        sheet.appendRow(newRow);
        console.log(`✓ 追加: ${course.name} (${course.id}) - ${mappedState}`);
        addCount++;
      }
    });

    console.log("\n" + "=".repeat(80));
    console.log("同期完了");
    console.log(`更新: ${updateCount}件, 追加: ${addCount}件`);
    console.log("=".repeat(80));

  } catch (error) {
    console.error("エラー:", error);
    throw error;
  }
}

/**
 * Classroom APIのcourseStateをクラスマスタの形式にマッピング
 * @param {string} apiState - Classroom APIのcourseState
 * @returns {string} クラスマスタの形式
 */
function mapCourseState(apiState) {
  const stateMap = {
    "ACTIVE": "Active",
    "ARCHIVED": "Archived",
    "PROVISIONED": "New",      // 作成されたが未公開
    "DECLINED": "Archived",    // 招待が拒否された
    "SUSPENDED": "Archived"    // 一時停止
  };

  return stateMap[apiState] || "Active";
}
