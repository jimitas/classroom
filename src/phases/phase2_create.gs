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
    }

    // 管理者アカウントIDを取得
    const adminAccountId = getAdminAccountId();
    console.log(`管理者アカウントID: ${adminAccountId}`);

    // クラス状態が "New" かつ 科目有効性が TRUE のクラスを取得
    const allClasses = getClassesByState(CONFIG.CLASS_STATE.NEW);
    const classesToCreate = allClasses.filter(c => c.isActive === true || c.isActive === "TRUE");

    if (classesToCreate.length === 0) {
      console.log("作成対象のクラスはありません");
      console.log("※ クラス状態が「New」かつ科目有効性が「TRUE」のクラスがありません");
      logPhaseComplete("Phase 2", successCount, errorCount);
      return { successCount, errorCount };
    }

    console.log(`作成対象クラス数: ${classesToCreate.length}`);

    // 各クラスに対して作成処理を実行
    for (const classInfo of classesToCreate) {
      try {
        createClass(classInfo, adminAccountId, dryRunMode);
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
 * @param {string} adminAccountId - 管理者アカウントID
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 */
function createClass(classInfo, adminAccountId, dryRunMode) {
  const subjectName = classInfo.subjectName;
  const sectionName = classInfo.section || ""; // セクション情報（オプション）

  debugLog(`クラス作成処理開始: ${subjectName}`);

  if (dryRunMode) {
    // DRY_RUNモード: ログのみ記録
    console.log(`[DRY-RUN] クラス作成: ${subjectName}`);
    console.log(`[DRY-RUN] オーナー: ${adminAccountId}`);

    // DRY-RUN用の仮クラスID
    const dummyCourseId = `DRYRUN_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    logToProcessLog({
      subjectName: subjectName,
      classroomId: dummyCourseId,
      result: "成功（DRY-RUN）",
      message: `クラス「${subjectName}」を作成予定（オーナー: ${adminAccountId}）`,
      executor: "Auto System (DRY-RUN)"
    });
  } else {
    // 実際のAPI呼び出し
    try {
      // クラスを作成
      const createdCourse = createClassroomCourse(subjectName, sectionName, adminAccountId);
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
        message: `クラス「${subjectName}」を作成しました（オーナー: ${adminAccountId}）`,
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
 * @param {string} ownerId - オーナーのメールアドレス
 * @returns {Object} 作成されたクラス情報
 */
function createClassroomCourse(subjectName, sectionName, ownerId) {
  return executeWithRetry(() => {
    const course = {
      name: subjectName,
      ownerId: ownerId,
      courseState: "ACTIVE"
    };

    // セクション名がある場合は追加
    if (sectionName) {
      course.section = sectionName;
    }

    return Classroom.Courses.create(course);
  });
}
