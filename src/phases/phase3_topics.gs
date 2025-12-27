/**
 * Phase 3: トピックの作成
 * @module Phase3Topics
 */

/**
 * Phase 3メイン処理: トピックの作成
 * @returns {Object} 実行結果 {successCount, errorCount}
 */
function runPhase3CreateTopics() {
  logPhaseStart("Phase 3: トピックの作成");

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

    console.log(`トピック構造: ${CONFIG.TOPIC_STRUCTURE.join(', ')}`);

    // クラス状態が "Active" かつ トピック作成フラグが TRUE のクラスを取得
    const allClasses = getClassesByState(CONFIG.CLASS_STATE.ACTIVE);

    // トピック作成フラグが TRUE のもののみをフィルタリング
    const classesForTopics = allClasses.filter(c => {
      return c.createTopics === true || c.createTopics === "TRUE";
    });

    // スキップされたクラスをログ出力
    const skippedClasses = allClasses.filter(c => {
      return c.createTopics === false || c.createTopics === "FALSE" || !c.createTopics;
    });

    if (skippedClasses.length > 0) {
      console.log(`スキップ: ${skippedClasses.length}件（トピック作成=FALSE）`);
      skippedClasses.forEach(c => {
        console.log(`  - ${c.subjectName}`);
      });
    }

    if (classesForTopics.length === 0) {
      console.log("作成対象のクラスはありません");
      console.log("※ クラス状態が「Active」かつトピック作成が「TRUE」のクラスがありません");
      logPhaseComplete("Phase 3", successCount, errorCount);
      return { successCount, errorCount };
    }

    console.log(`作成対象クラス数: ${classesForTopics.length}`);

    // 各クラスに対してトピック作成処理を実行
    for (const classInfo of classesForTopics) {
      try {
        const result = createTopicsForClass(classInfo, dryRunMode);

        if (result.skipped) {
          console.log(`✓ スキップ: ${classInfo.subjectName} - 既に全トピックが存在`);
        } else {
          successCount++;
        }

        // API Rate Limit対策: リクエスト間隔を空ける
        if (!dryRunMode && !result.skipped) {
          Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
        }
      } catch (error) {
        console.error(`トピック作成エラー: ${classInfo.subjectName}`, error);
        errorCount++;

        // エラー処理
        handleError(error, {
          subjectName: classInfo.subjectName,
          classroomId: classInfo.classroomId
        });
      }
    }

    logPhaseComplete("Phase 3", successCount, errorCount);
  } catch (error) {
    console.error("Phase 3実行エラー:", error);
    throw error;
  }

  return { successCount, errorCount };
}

/**
 * 個別クラスのトピック作成処理
 * @param {Object} classInfo - クラス情報
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 * @returns {Object} 実行結果 {skipped: boolean}
 */
function createTopicsForClass(classInfo, dryRunMode) {
  const subjectName = classInfo.subjectName;
  const courseId = classInfo.classroomId;

  debugLog(`トピック作成処理開始: ${subjectName}`);

  // 既存トピックを取得して重複チェック
  const existingTopics = getExistingTopics(courseId, dryRunMode);
  const existingTopicNames = existingTopics.map(t => t.name);

  // 作成対象のトピック（既存にないもののみ）
  const topicsToCreate = CONFIG.TOPIC_STRUCTURE.filter(topicName => {
    return !existingTopicNames.includes(topicName);
  });

  if (topicsToCreate.length === 0) {
    // 全トピックが既に存在する場合はスキップ
    logToProcessLog({
      subjectName: subjectName,
      classroomId: courseId,
      result: "スキップ",
      message: `全トピックが既に存在します（${CONFIG.TOPIC_STRUCTURE.length}件）`,
      executor: dryRunMode ? "Auto System (DRY-RUN)" : "Auto System"
    });

    return { skipped: true };
  }

  debugLog(`作成するトピック: ${topicsToCreate.join(', ')}`);

  if (dryRunMode) {
    // DRY_RUNモード: ログのみ記録
    console.log(`[DRY-RUN] トピック作成: ${subjectName}`);
    console.log(`[DRY-RUN] 作成数: ${topicsToCreate.length}件 / ${CONFIG.TOPIC_STRUCTURE.length}件`);

    logToProcessLog({
      subjectName: subjectName,
      classroomId: courseId,
      result: "成功（DRY-RUN）",
      message: `トピック${topicsToCreate.length}件を作成予定: ${topicsToCreate.join(', ')}`,
      executor: "Auto System (DRY-RUN)"
    });
  } else {
    // 実際のAPI呼び出し
    try {
      let createdCount = 0;

      // トピックを作成（逆順で作成されているので、そのまま順番に実行）
      for (const topicName of topicsToCreate) {
        createClassroomTopic(courseId, topicName);
        createdCount++;
        debugLog(`トピック作成完了: ${topicName}`);

        // トピック作成間隔を空ける
        if (createdCount < topicsToCreate.length) {
          Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
        }
      }

      // 成功ログを記録
      logToProcessLog({
        subjectName: subjectName,
        classroomId: courseId,
        result: "成功",
        message: `トピック${createdCount}件を作成しました: ${topicsToCreate.join(', ')}`,
        executor: "Auto System"
      });

      console.log(`✓ トピック作成完了: ${subjectName} (${createdCount}件)`);
    } catch (error) {
      // エラーログを記録
      logToProcessLog({
        subjectName: subjectName,
        classroomId: courseId,
        result: "失敗",
        message: error.message,
        executor: "Auto System"
      });

      throw error;
    }
  }

  return { skipped: false };
}

/**
 * 既存トピックを取得
 * @param {string} courseId - クラスID
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 * @returns {Array<Object>} 既存トピックの配列
 */
function getExistingTopics(courseId, dryRunMode) {
  if (dryRunMode) {
    // DRY_RUNモードでは空配列を返す（重複チェックをスキップ）
    return [];
  }

  return executeWithRetry(() => {
    const response = Classroom.Courses.Topics.list(courseId);
    return response.topic || [];
  });
}

/**
 * Classroom APIを使用してトピックを作成
 * @param {string} courseId - クラスID
 * @param {string} topicName - トピック名
 * @returns {Object} 作成されたトピック情報
 */
function createClassroomTopic(courseId, topicName) {
  return executeWithRetry(() => {
    const topic = {
      name: topicName
    };

    return Classroom.Courses.Topics.create(topic, courseId);
  });
}
