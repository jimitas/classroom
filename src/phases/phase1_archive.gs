/**
 * Phase 1: 旧年度クラスのアーカイブ
 * @module Phase1Archive
 */

/**
 * Phase 1メイン処理: 旧年度クラスのアーカイブ
 * @returns {Object} 実行結果 {successCount, errorCount}
 */
function runPhase1Archive() {
  logPhaseStart("Phase 1: 旧年度クラスのアーカイブ");

  let successCount = 0;
  let errorCount = 0;

  try {
    // DRY_RUN_MODEを確認
    const dryRunMode = isDryRunMode();
    if (dryRunMode) {
      console.log("⚠️ DRY_RUN_MODE: 実際のAPI呼び出しは行いません");
    }

    // クラス状態が "Archived" のクラスを取得
    const classesToArchive = getClassesByState(CONFIG.CLASS_STATE.ARCHIVED);

    if (classesToArchive.length === 0) {
      console.log("アーカイブ対象のクラスはありません");
      logPhaseComplete("Phase 1", successCount, errorCount);
      return { successCount, errorCount };
    }

    console.log(`アーカイブ対象クラス数: ${classesToArchive.length}`);

    // 各クラスに対してアーカイブ処理を実行
    for (const classInfo of classesToArchive) {
      try {
        archiveClass(classInfo, dryRunMode);
        successCount++;

        // API Rate Limit対策: リクエスト間隔を空ける
        if (!dryRunMode) {
          Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
        }
      } catch (error) {
        console.error(`クラスアーカイブエラー: ${classInfo.subjectName}`, error);
        errorCount++;

        // エラー処理
        handleError(error, {
          subjectName: classInfo.subjectName,
          classroomId: classInfo.classroomId
        });
      }
    }

    logPhaseComplete("Phase 1", successCount, errorCount);
  } catch (error) {
    console.error("Phase 1実行エラー:", error);
    throw error;
  }

  return { successCount, errorCount };
}

/**
 * 個別クラスのアーカイブ処理
 * @param {Object} classInfo - クラス情報
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 */
function archiveClass(classInfo, dryRunMode) {
  const classroomId = classInfo.classroomId;
  const originalName = classInfo.subjectName;

  // 年度プレフィックスを生成（現在の年度を使用）
  const currentYear = new Date().getFullYear();
  const archivedName = `${currentYear}_${originalName}`;

  debugLog(`アーカイブ処理開始: ${originalName} → ${archivedName}`);

  if (dryRunMode) {
    // DRY_RUNモード: ログのみ記録
    console.log(`[DRY-RUN] クラス名変更: ${originalName} → ${archivedName}`);
    console.log(`[DRY-RUN] アーカイブ実行: ${classroomId}`);

    logToProcessLog({
      subjectName: originalName,
      classroomId: classroomId,
      result: "成功（DRY-RUN）",
      message: `クラス名を「${archivedName}」に変更し、アーカイブ予定`,
      executor: "Auto System (DRY-RUN)"
    });
  } else {
    // 実際のAPI呼び出し
    try {
      // クラス名変更とアーカイブを1回のAPI呼び出しで実行
      updateAndArchiveClassroom(classroomId, archivedName);
      debugLog(`クラス名変更とアーカイブ完了: ${archivedName}`);

      // 成功ログを記録
      logToProcessLog({
        subjectName: originalName,
        classroomId: classroomId,
        result: "成功",
        message: `クラス名を「${archivedName}」に変更し、アーカイブしました`,
        executor: "Auto System"
      });

      console.log(`✓ アーカイブ完了: ${originalName}`);
    } catch (error) {
      // エラーログを記録
      logToProcessLog({
        subjectName: originalName,
        classroomId: classroomId,
        result: "失敗",
        message: error.message,
        executor: "Auto System"
      });

      throw error;
    }
  }
}

/**
 * Classroom APIを使用してクラス名を変更してアーカイブ
 * @param {string} courseId - クラスID
 * @param {string} newName - 新しいクラス名
 */
function updateAndArchiveClassroom(courseId, newName) {
  return executeWithRetry(() => {
    // クラス名とアーカイブ状態を同時に更新
    // Google Classroom APIはcourses.updateで部分更新ができないため、
    // 両方を同時に指定する必要がある
    // 注意: リクエストボディにidを含めてはいけない（第2引数で渡す）
    const course = {
      name: newName,
      courseState: "ARCHIVED"
    };

    return Classroom.Courses.update(course, courseId);
  });
}
