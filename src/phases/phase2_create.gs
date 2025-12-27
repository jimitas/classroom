/**
 * Phase 2: 新年度クラスの作成
 * @module Phase2Create
 */

/**
 * Phase 2メイン処理: 新年度クラスの作成
 * @returns {Object} 実行結果 {successCount, errorCount}
 */
function runPhase2CreateClasses() {
  logPhaseStart("Phase 2: 新年度クラスの作成");

  let successCount = 0;
  let errorCount = 0;

  try {
    // DRY_RUN_MODEを確認
    const dryRunMode = isDryRunMode();
    if (dryRunMode) {
      console.log("⚠️ DRY_RUN_MODE: 実際のAPI呼び出しは行いません");
      console.log("⚠️ 本番実行するには、「システム設定 」シートでDRY_RUN_MODEをFALSEに設定してください");
      SpreadsheetApp.getActiveSpreadsheet().toast('DRY_RUNモードで実行中です', '警告', 5);
    } else {
      console.log("✅ 本番モード: 実際にAPI呼び出しを行います");
      SpreadsheetApp.getActiveSpreadsheet().toast('本番モードで実行中...', '実行中', 3);
    }

    console.log("クラスのオーナー: スクリプト実行者");

    // クラス状態が "New" かつ 科目有効性が TRUE のクラスを取得
    const allClasses = getClassesByState(CONFIG.CLASS_STATE.NEW);

    // クラスIDが空のもののみをフィルタリング（重複作成防止）
    const classesToCreate = allClasses.filter(c => {
      const isActive = c.isActive === true || c.isActive === "TRUE";
      const hasNoClassroomId = !c.classroomId || c.classroomId === "";
      return isActive && hasNoClassroomId;
    });

    // スキップされたクラスをログ出力
    const skippedClasses = allClasses.filter(c => {
      const isActive = c.isActive === true || c.isActive === "TRUE";
      const hasClassroomId = c.classroomId && c.classroomId !== "";
      return isActive && hasClassroomId;
    });

    if (skippedClasses.length > 0) {
      console.log(`スキップ: ${skippedClasses.length}件（既にクラスIDが存在）`);
      skippedClasses.forEach(c => {
        console.log(`  - ${c.subjectName} (${c.classroomId})`);
      });
    }

    if (classesToCreate.length === 0) {
      console.log("作成対象のクラスはありません");
      console.log("※ クラス状態が「New」かつ科目有効性が「TRUE」かつクラスIDが空のクラスがありません");
      logPhaseComplete("Phase 2", successCount, errorCount);
      return { successCount, errorCount };
    }

    console.log(`作成対象クラス数: ${classesToCreate.length}`);

    // 各クラスに対して作成処理を実行
    for (const classInfo of classesToCreate) {
      try {
        createClass(classInfo, dryRunMode);
        successCount++;

        // API Rate Limit対策: リクエスト間隔を空ける
        if (!dryRunMode) {
          Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
        }
      } catch (error) {
        console.error(`クラス作成エラー: ${classInfo.subjectName}`, error);
        errorCount++;

        // エラー処理
        handleError(error, {
          subjectName: classInfo.subjectName,
          classroomId: classInfo.classroomId
        });
      }
    }

    logPhaseComplete("Phase 2", successCount, errorCount);
  } catch (error) {
    console.error("Phase 2実行エラー:", error);
    throw error;
  }

  return { successCount, errorCount };
}

/**
 * 個別クラスの作成処理
 * @param {Object} classInfo - クラス情報
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 */
function createClass(classInfo, dryRunMode) {
  const subjectName = classInfo.subjectName;
  const sectionName = classInfo.section || ""; // セクション情報（オプション）

  debugLog(`クラス作成処理開始: ${subjectName}`);

  if (dryRunMode) {
    // DRY_RUNモード: ログのみ記録
    console.log(`[DRY-RUN] クラス作成: ${subjectName}`);
    console.log(`[DRY-RUN] オーナー: スクリプト実行者`);

    // DRY-RUN用の仮クラスID
    const dummyCourseId = `DRYRUN_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    logToProcessLog({
      subjectName: subjectName,
      classroomId: dummyCourseId,
      result: "成功（DRY-RUN）",
      message: `クラス「${subjectName}」を作成予定（オーナー: スクリプト実行者）`,
      executor: "Auto System (DRY-RUN)"
    });
  } else {
    // 実際のAPI呼び出し
    try {
      // クラスを作成（オーナー: スクリプト実行者）
      const createdCourse = createClassroomCourse(subjectName, sectionName);
      const courseId = createdCourse.id;
      debugLog(`クラス作成完了: ${subjectName} (${courseId})`);

      // クラスマスタを更新（クラスIDと状態を更新）
      updateClassMaster(classInfo.rowIndex, {
        classroomId: courseId,
        state: CONFIG.CLASS_STATE.ACTIVE
      });
      debugLog(`クラスマスタ更新完了: ${subjectName}`);

      // 成功ログを記録
      logToProcessLog({
        subjectName: subjectName,
        classroomId: courseId,
        result: "成功",
        message: `クラス「${subjectName}」を作成しました（オーナー: スクリプト実行者）`,
        executor: "Auto System"
      });

      console.log(`✓ クラス作成完了: ${subjectName} (${courseId})`);
    } catch (error) {
      // エラーログを記録
      logToProcessLog({
        subjectName: subjectName,
        classroomId: "",
        result: "失敗",
        message: error.message,
        executor: "Auto System"
      });

      throw error;
    }
  }
}

/**
 * Classroom APIを使用してクラスを作成
 * @param {string} subjectName - 科目名
 * @param {string} sectionName - セクション名（オプション）
 * @returns {Object} 作成されたクラス情報
 */
function createClassroomCourse(subjectName, sectionName) {
  return executeWithRetry(() => {
    const course = {
      name: subjectName,
      ownerId: 'me',  // 'me' = スクリプト実行者
      courseState: "ACTIVE"
    };

    // セクション名がある場合は追加
    if (sectionName) {
      course.section = sectionName;
    }

    return Classroom.Courses.create(course);
  });
}
