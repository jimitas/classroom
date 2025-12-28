/**
 * Phase 5: 招待状況の確認
 * @module Phase5InvitationStatus
 */

/**
 * Phase 5メイン処理: 招待状況の確認
 * @returns {Object} 実行結果 {totalInvitations, pendingCount, acceptedCount}
 */
function runPhase5CheckInvitationStatus() {
  logPhaseStart("Phase 5: 招待状況の確認");

  let totalInvitations = 0;
  let pendingCount = 0;
  let acceptedCount = 0;

  try {
    // Active状態のクラスを取得
    const allClasses = getClassesByState(CONFIG.CLASS_STATE.ACTIVE);
    const activeClasses = allClasses.filter(c => c.isActive === true || c.isActive === "TRUE");

    if (activeClasses.length === 0) {
      console.log("確認対象のクラスはありません");
      logPhaseComplete("Phase 5", 0, 0);
      return { totalInvitations: 0, pendingCount: 0, acceptedCount: 0 };
    }

    console.log(`確認対象クラス数: ${activeClasses.length}`);

    // 各クラスの招待状況を確認
    for (const classInfo of activeClasses) {
      try {
        const result = checkClassInvitations(classInfo);

        totalInvitations += result.pendingInvitations.length;
        pendingCount += result.pendingInvitations.length;

        // API Rate Limit対策
        Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
      } catch (error) {
        console.error(`招待状況確認エラー: ${classInfo.subjectName}`, error);
      }
    }

    // 承諾済み数を計算（仮: 総招待数から保留中を引く）
    // 実際には Students/Teachers リストから取得可能
    acceptedCount = totalInvitations > 0 ? totalInvitations - pendingCount : 0;

    logPhaseComplete("Phase 5", activeClasses.length, 0);

    console.log(`\n招待状況サマリー:`);
    console.log(`  保留中: ${pendingCount}件`);
    console.log(`  確認済みクラス: ${activeClasses.length}件`);

  } catch (error) {
    console.error("Phase 5実行エラー:", error);
    throw error;
  }

  return { totalInvitations, pendingCount, acceptedCount };
}

/**
 * 個別クラスの招待状況を確認
 * @param {Object} classInfo - クラス情報
 * @returns {Object} 招待状況 {pendingInvitations}
 */
function checkClassInvitations(classInfo) {
  const courseId = classInfo.classroomId;
  const subjectName = classInfo.subjectName;

  debugLog(`招待状況確認開始: ${subjectName} (${courseId})`);

  try {
    // 保留中の招待を取得
    const pendingInvitations = getPendingInvitations(courseId);

    if (pendingInvitations.length === 0) {
      console.log(`✓ ${subjectName}: 保留中の招待なし`);

      logToProcessLog({
        subjectName: subjectName,
        classroomId: courseId,
        result: "確認完了",
        message: "保留中の招待はありません",
        executor: "Auto System"
      });
    } else {
      console.log(`⚠️ ${subjectName}: 保留中の招待 ${pendingInvitations.length}件`);

      // 保留中の招待をログに記録
      pendingInvitations.forEach(invitation => {
        const role = invitation.role === "STUDENT" ? "生徒" : "教員";
        const userId = invitation.userId || "不明";

        logToProcessLog({
          subjectName: subjectName,
          classroomId: courseId,
          studentName: userId,
          result: "保留中",
          message: `招待が未承諾です（${role}）`,
          executor: "Auto System"
        });

        debugLog(`  - ${role}: ${userId}`);
      });
    }

    return { pendingInvitations };
  } catch (error) {
    console.error(`招待状況確認エラー: ${subjectName}`, error);

    logToProcessLog({
      subjectName: subjectName,
      classroomId: courseId,
      result: "失敗",
      message: `招待状況確認エラー: ${error.message}`,
      executor: "Auto System"
    });

    throw error;
  }
}

/**
 * クラスの全メンバー（承諾済み）を取得
 * @param {string} courseId - クラスID
 * @returns {Object} メンバー情報 {students, teachers}
 */
function getAcceptedMembers(courseId) {
  const students = executeWithRetry(() => {
    try {
      const response = Classroom.Courses.Students.list(courseId);
      return response.students || [];
    } catch (error) {
      if (error.message && error.message.includes("not found")) {
        return [];
      }
      throw error;
    }
  });

  const teachers = executeWithRetry(() => {
    try {
      const response = Classroom.Courses.Teachers.list(courseId);
      return response.teachers || [];
    } catch (error) {
      if (error.message && error.message.includes("not found")) {
        return [];
      }
      throw error;
    }
  });

  return { students, teachers };
}

/**
 * 招待状況の詳細レポートを生成（デバッグ用）
 * @param {string} courseId - クラスID（省略時は全Activeクラス）
 */
function generateInvitationReport(courseId) {
  console.log("=".repeat(60));
  console.log("招待状況詳細レポート");
  console.log("=".repeat(60));

  const classes = courseId
    ? [{ classroomId: courseId, subjectName: "指定クラス" }]
    : getClassesByState(CONFIG.CLASS_STATE.ACTIVE);

  classes.forEach(classInfo => {
    console.log(`\nクラス: ${classInfo.subjectName} (${classInfo.classroomId})`);
    console.log("-".repeat(60));

    const pendingInvitations = getPendingInvitations(classInfo.classroomId);
    const acceptedMembers = getAcceptedMembers(classInfo.classroomId);

    console.log(`保留中の招待: ${pendingInvitations.length}件`);
    pendingInvitations.forEach(inv => {
      console.log(`  - ${inv.role}: ${inv.userId}`);
    });

    console.log(`\n承諾済み生徒: ${acceptedMembers.students.length}名`);
    console.log(`承諾済み教員: ${acceptedMembers.teachers.length}名`);
  });

  console.log("\n" + "=".repeat(60));
}
