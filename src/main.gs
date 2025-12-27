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

    // Phase 2: クラス作成
    results.phase2 = runPhase2CreateClasses();

    // Phase 3: トピック作成
    results.phase3 = runPhase3CreateTopics();

    // Phase 4: 生徒・教員の一括登録
    results.phase4 = runPhase4RegisterMembers();

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
 * Phase 2のみを実行（テスト用）
 */
function testPhase2() {
  console.log("Phase 2テスト実行");
  const result = runPhase2CreateClasses();
  console.log("Phase 2テスト完了:", result);
}

/**
 * Phase 3のみを実行（テスト用）
 */
function testPhase3() {
  console.log("Phase 3テスト実行");
  const result = runPhase3CreateTopics();
  console.log("Phase 3テスト完了:", result);
}

/**
 * Phase 4のみを実行（テスト用）
 */
function testPhase4() {
  console.log("Phase 4テスト実行");
  const result = runPhase4RegisterMembers();
  console.log("Phase 4テスト完了:", result);
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

/**
 * クラスのオーナー情報を確認（デバッグ用）
 * @param {string} courseId - クラスID（省略時は"823999925372"）
 */
function checkClassOwner(courseId) {
  if (!courseId) {
    courseId = "823999925372";  // デフォルトのクラスID
  }

  try {
    const course = Classroom.Courses.get(courseId);

    // 現在のユーザー情報を取得（Classroom API使用）
    const currentUser = Classroom.UserProfiles.get('me');
    const executorEmail = currentUser.emailAddress;

    console.log("=".repeat(60));
    console.log("クラス情報");
    console.log("=".repeat(60));
    console.log("クラス名:", course.name);
    console.log("クラスID:", course.id);
    console.log("オーナーID:", course.ownerId);
    console.log("クラス状態:", course.courseState);
    console.log("\nスクリプト実行者:", executorEmail);
    console.log("実行者ID:", currentUser.id);

    // 教師一覧を取得
    console.log("\n" + "=".repeat(60));
    console.log("教師一覧");
    console.log("=".repeat(60));
    const teachers = Classroom.Courses.Teachers.list(courseId);
    if (teachers.teachers && teachers.teachers.length > 0) {
      teachers.teachers.forEach(t => {
        const isOwner = t.userId === course.ownerId ? " ★オーナー" : "";
        const isExecutor = t.userId === currentUser.id ? " (あなた)" : "";
        console.log(`  - ${t.profile.name.fullName} (${t.profile.emailAddress})${isOwner}${isExecutor}`);
      });
    } else {
      console.log("  教師が見つかりません");
    }

    // スクリプト実行者が教師かどうか確認
    const isTeacher = teachers.teachers && teachers.teachers.some(t =>
      t.userId === currentUser.id
    );
    const isOwner = course.ownerId === currentUser.id;

    console.log("\n" + "=".repeat(60));
    console.log("権限チェック");
    console.log("=".repeat(60));
    console.log("スクリプト実行者は教師:", isTeacher ? "はい ✅" : "いいえ ❌");
    console.log("スクリプト実行者はオーナー:", isOwner ? "はい ✅" : "いいえ ❌");

    if (!isTeacher) {
      console.log("\n⚠️ このクラスに対する権限がありません");
      console.log("対処方法: Classroomでこのクラスに共同教師として追加してください");
    } else if (!isOwner) {
      console.log("\n⚠️ 教員を追加するにはオーナー権限が必要です");
      console.log("対処方法: Classroomでオーナーを移譲してもらってください");
    } else {
      console.log("\n✅ 生徒・教員の追加が可能です");
    }

  } catch (error) {
    console.error("エラー:", error);
    console.error("クラスへのアクセス権限がない可能性があります");
  }
}

/**
 * Phase 4デバッグ用（一時的な関数）
 */
function debugPhase4Data() {
  console.log("=== Phase 4 デバッグ ===");

  // 1. Active クラスを取得
  const allClasses = getClassesByState(CONFIG.CLASS_STATE.ACTIVE);
  const activeClasses = allClasses.filter(c => c.isActive === true || c.isActive === "TRUE");

  console.log(`\nActive クラス数: ${activeClasses.length}`);
  activeClasses.forEach(c => {
    console.log(`  - 科目名: "${c.subjectName}", 履修列名: "${c.enrollmentColumn}", ID: ${c.classroomId}`);
  });

  // 2. データ変換
  const classMaps = createClassMasterMaps(activeClasses);

  console.log(`\n履修列マップ:`, Object.keys(classMaps.byEnrollmentColumn));
  console.log(`科目名マップ:`, Object.keys(classMaps.bySubjectName));

  const enrollmentData = getEnrollmentMaster();
  const teacherData = getTeacherMaster();

  const studentEnrollments = transformEnrollmentDataToVertical(enrollmentData, classMaps.byEnrollmentColumn);
  const teacherAssignments = transformTeacherDataToVertical(teacherData, classMaps.bySubjectName);

  console.log(`\n生徒履修データ: ${studentEnrollments.length}件`);
  studentEnrollments.forEach(e => {
    console.log(`  - ${e.studentName} (${e.studentId}) → "${e.subjectName}" (${e.classroomId || "IDなし"})`);
  });

  console.log(`\n教員担当データ: ${teacherAssignments.length}件`);
  teacherAssignments.forEach(t => {
    console.log(`  - ${t.teacherName} → "${t.subjectName}" (${t.classroomId || "IDなし"})`);
  });

  // 3. 履修登録マスタのヘッダー確認
  console.log(`\n履修登録マスタ ヘッダー（3列目以降）:`);
  const headers = enrollmentData[0];
  for (let i = 2; i < headers.length; i++) {
    console.log(`  列${i}: "${headers[i]}"`);
  }

  // 4. 教員マスタのヘッダー確認
  console.log(`\n教員マスタ ヘッダー（4列目以降）:`);
  const teacherHeaders = teacherData[0];
  for (let i = 3; i < teacherHeaders.length; i++) {
    console.log(`  列${i}: "${teacherHeaders[i]}"`);
  }
}

/**
 * API詳細診断: 1人の生徒を追加してエラー詳細を確認
 */
function diagnoseApiPermission() {
  const courseId = "823999925372";  // テストクラスID
  const testStudentEmail = "tlu-test01@shs.kyoto-art.ac.jp";  // テスト生徒

  console.log("=".repeat(60));
  console.log("Classroom API 詳細診断");
  console.log("=".repeat(60));

  try {
    // 1. クラス情報を取得
    const course = Classroom.Courses.get(courseId);
    console.log("クラス名:", course.name);
    console.log("クラス状態:", course.courseState);
    console.log("オーナーID:", course.ownerId);

    // 2. 現在のユーザー情報
    const currentUser = Classroom.UserProfiles.get('me');
    console.log("\n実行者:", currentUser.emailAddress);
    console.log("実行者ID:", currentUser.id);
    console.log("実行者はオーナー:", course.ownerId === currentUser.id ? "はい" : "いいえ");

    // 3. 既存の生徒を確認
    console.log("\n" + "=".repeat(60));
    console.log("既存の生徒リスト");
    console.log("=".repeat(60));
    const existingStudents = Classroom.Courses.Students.list(courseId);
    if (existingStudents.students && existingStudents.students.length > 0) {
      existingStudents.students.forEach(s => {
        console.log(`  - ${s.profile.name.fullName} (${s.profile.emailAddress})`);
      });
    } else {
      console.log("  （生徒なし）");
    }

    // 4. テスト: 生徒を追加してみる
    console.log("\n" + "=".repeat(60));
    console.log("生徒追加テスト");
    console.log("=".repeat(60));
    console.log("追加対象:", testStudentEmail);

    try {
      const student = {
        userId: testStudentEmail
      };
      const result = Classroom.Courses.Students.create(student, courseId);
      console.log("✅ 成功！生徒を追加できました");
      console.log("追加された生徒:", result.profile.name.fullName);
    } catch (error) {
      console.error("❌ 失敗！エラー詳細:");
      console.error("エラーメッセージ:", error.message);
      console.error("エラー詳細:", JSON.stringify(error, null, 2));

      // 考えられる原因を出力
      console.log("\n考えられる原因:");
      if (error.message.includes("does not have permission")) {
        console.log("1. Google Workspace管理コンソールでAPIアクセスが制限されている");
        console.log("2. ドメイン間での生徒追加がブロックされている");
        console.log("3. サービスアカウントとドメイン全体の委任が必要");
      }
    }

  } catch (error) {
    console.error("診断中にエラーが発生:", error);
  }
}
