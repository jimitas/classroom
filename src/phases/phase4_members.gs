/**
 * Phase 4: 生徒・教員の一括登録
 * @module Phase4Members
 */

/**
 * Phase 4メイン処理: 生徒・教員の一括登録
 * @returns {Object} 実行結果 {successCount, errorCount, studentCount, teacherCount}
 */
function runPhase4RegisterMembers() {
  logPhaseStart("Phase 4: 生徒・教員の一括登録");

  let successCount = 0;
  let errorCount = 0;
  let studentRegistered = 0;
  let teacherRegistered = 0;

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

    console.log("=".repeat(60));
    console.log("データ読み込み開始");
    console.log("=".repeat(60));

    // 1. クラスマスタから Active かつ 科目有効性=TRUE のクラスを取得
    const allClasses = getClassesByState(CONFIG.CLASS_STATE.ACTIVE);
    const activeClasses = allClasses.filter(c => c.isActive === true || c.isActive === "TRUE");

    if (activeClasses.length === 0) {
      console.log("登録対象のクラスはありません");
      console.log("※ クラス状態が「Active」かつ科目有効性が「TRUE」のクラスがありません");
      logPhaseComplete("Phase 4", successCount, errorCount);
      return { successCount, errorCount, studentCount: 0, teacherCount: 0 };
    }

    console.log(`登録対象クラス数: ${activeClasses.length}`);

    // 2. クラスマスタからマップを生成（検索用）
    const classMaps = createClassMasterMaps(activeClasses);

    // 3. 履修登録マスタを読み込み
    const enrollmentData = getEnrollmentMaster();
    console.log(`履修登録マスタ: ${enrollmentData.length - 1}行読み込み`);

    // 4. 教員マスタを読み込み
    const teacherData = getTeacherMaster();
    console.log(`教員マスタ: ${teacherData.length - 1}行読み込み`);

    // 5. アカウントマッピングを読み込み
    const accountMap = getAccountMappingMap();
    console.log(`アカウントマッピング: ${Object.keys(accountMap).length}件読み込み`);

    console.log("=".repeat(60));
    console.log("データ変換開始");
    console.log("=".repeat(60));

    // 6. マトリクス→縦持ち変換（生徒）
    const studentEnrollments = transformEnrollmentDataToVertical(
      enrollmentData,
      classMaps.byEnrollmentColumn
    );
    console.log(`生徒履修データ: ${studentEnrollments.length}件に変換`);

    // 7. マトリクス→縦持ち変換（教員）
    const teacherAssignments = transformTeacherDataToVertical(
      teacherData,
      classMaps.bySubjectName
    );
    console.log(`教員担当データ: ${teacherAssignments.length}件に変換`);

    console.log("=".repeat(60));
    console.log("登録処理開始");
    console.log("=".repeat(60));

    // 8. クラスごとにグループ化して登録
    const classSummary = registerMembersByClass(
      studentEnrollments,
      teacherAssignments,
      accountMap,
      dryRunMode
    );

    successCount = classSummary.successCount;
    errorCount = classSummary.errorCount;
    studentRegistered = classSummary.studentCount;
    teacherRegistered = classSummary.teacherCount;

    logPhaseComplete("Phase 4", successCount, errorCount);

  } catch (error) {
    console.error("Phase 4実行エラー:", error);
    throw error;
  }

  return {
    successCount,
    errorCount,
    studentCount: studentRegistered,
    teacherCount: teacherRegistered
  };
}

/**
 * クラスごとにメンバーを登録
 * @param {Array} studentEnrollments - 生徒履修データ（縦持ち）
 * @param {Array} teacherAssignments - 教員担当データ（縦持ち）
 * @param {Object} accountMap - アカウントマッピング
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 * @returns {Object} 実行結果
 */
function registerMembersByClass(studentEnrollments, teacherAssignments, accountMap, dryRunMode) {
  let successCount = 0;
  let errorCount = 0;
  let totalStudents = 0;
  let totalTeachers = 0;

  // クラスIDごとにグループ化
  const classSummary = {};

  // 生徒をグループ化
  studentEnrollments.forEach(enrollment => {
    const classId = enrollment.classroomId;
    if (!classSummary[classId]) {
      classSummary[classId] = {
        subjectName: enrollment.subjectName,
        students: [],
        teachers: []
      };
    }
    classSummary[classId].students.push(enrollment);
  });

  // 教員をグループ化
  teacherAssignments.forEach(assignment => {
    const classId = assignment.classroomId;
    if (!classSummary[classId]) {
      classSummary[classId] = {
        subjectName: assignment.subjectName,
        students: [],
        teachers: []
      };
    }
    classSummary[classId].teachers.push(assignment);
  });

  // クラスごとに登録処理
  for (const classId in classSummary) {
    const classData = classSummary[classId];

    try {
      const result = registerClassMembers(
        classId,
        classData.subjectName,
        classData.students,
        classData.teachers,
        accountMap,
        dryRunMode
      );

      if (!result.skipped) {
        successCount++;
        totalStudents += result.studentCount;
        totalTeachers += result.teacherCount;
      }

      // API Rate Limit対策: クラス間のみ待機（同一クラス内のメンバーは連続送信）
      if (!dryRunMode && !result.skipped) {
        Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
      }

    } catch (error) {
      console.error(`メンバー登録エラー: ${classData.subjectName} (${classId})`, error);
      errorCount++;

      handleError(error, {
        subjectName: classData.subjectName,
        classroomId: classId
      });
    }
  }

  return {
    successCount,
    errorCount,
    studentCount: totalStudents,
    teacherCount: totalTeachers
  };
}

/**
 * 個別クラスのメンバー登録処理
 * @param {string} classId - クラスID
 * @param {string} subjectName - 科目名
 * @param {Array} students - 生徒リスト
 * @param {Array} teachers - 教員リスト
 * @param {Object} accountMap - アカウントマッピング
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 * @returns {Object} 実行結果
 */
function registerClassMembers(classId, subjectName, students, teachers, accountMap, dryRunMode) {
  debugLog(`メンバー登録開始: ${subjectName} (${classId})`);

  let studentCount = 0;
  let teacherCount = 0;
  const errors = [];

  // 生徒のバリデーション
  students.forEach(student => {
    const validationError = validateStudentAccount(student, accountMap);
    if (validationError) {
      errors.push(validationError);
      return;
    }

    if (!dryRunMode) {
      try {
        addStudentToClass(classId, accountMap[student.studentId].email);
        studentCount++;
        debugLog(`生徒招待送信成功: ${student.studentName} → ${subjectName}`);
      } catch (error) {
        errors.push(`生徒「${student.studentName}」: ${error.message}`);
      }
    } else {
      studentCount++;
    }
  });

  // 教員のバリデーション
  teachers.forEach(teacher => {
    const validationError = validateTeacherAccount(teacher);
    if (validationError) {
      errors.push(validationError);
      return;
    }

    if (!dryRunMode) {
      try {
        addTeacherToClass(classId, teacher.teacherEmail);
        teacherCount++;
        debugLog(`教員招待送信成功: ${teacher.teacherName} → ${subjectName}`);
      } catch (error) {
        errors.push(`教員「${teacher.teacherName}」: ${error.message}`);
      }
    } else {
      teacherCount++;
    }
  });

  // ログ記録
  const modeLabel = dryRunMode ? "（DRY-RUN）" : "";
  const actionLabel = dryRunMode ? "招待送信予定" : "に招待を送信しました";

  const message = errors.length > 0
    ? `生徒${studentCount}名、教員${teacherCount}名${actionLabel}。エラー${errors.length}件: ${errors.join('; ')}`
    : `生徒${studentCount}名、教員${teacherCount}名${actionLabel}`;

  const result = errors.length > 0
    ? `一部失敗${modeLabel}`
    : `成功${modeLabel}`;

  logToProcessLog({
    subjectName: subjectName,
    classroomId: classId,
    result: result,
    message: message,
    executor: dryRunMode ? "Auto System (DRY-RUN)" : "Auto System"
  });

  if (dryRunMode) {
    console.log(`[DRY-RUN] ${subjectName}: 生徒${studentCount}名, 教員${teacherCount}名に招待送信予定 (エラー: ${errors.length}件)`);
  } else {
    console.log(`✓ ${subjectName}: 生徒${studentCount}名, 教員${teacherCount}名に招待送信`);
  }

  return { skipped: false, studentCount, teacherCount };
}

/**
 * 生徒をクラスに招待（Invitations API使用）
 * @param {string} courseId - クラスID
 * @param {string} studentEmail - 生徒のメールアドレス
 * @returns {Object} 招待結果
 */
function addStudentToClass(courseId, studentEmail) {
  return executeWithRetry(() => {
    const invitation = {
      userId: studentEmail,
      courseId: courseId,
      role: "STUDENT"
    };
    return Classroom.Invitations.create(invitation);
  });
}

/**
 * 教員をクラスに招待（共同教師として、Invitations API使用）
 * @param {string} courseId - クラスID
 * @param {string} teacherEmail - 教員のメールアドレス
 * @returns {Object} 招待結果
 */
function addTeacherToClass(courseId, teacherEmail) {
  return executeWithRetry(() => {
    const invitation = {
      userId: teacherEmail,
      courseId: courseId,
      role: "TEACHER"
    };
    return Classroom.Invitations.create(invitation);
  });
}

/**
 * 生徒アカウントのバリデーション
 * @param {Object} student - 生徒情報
 * @param {Object} accountMap - アカウントマッピング
 * @returns {string|null} エラーメッセージ、または null（正常）
 */
function validateStudentAccount(student, accountMap) {
  // アカウント存在チェック
  const account = accountMap[student.studentId];
  if (!account) {
    return `生徒「${student.studentName}」(${student.studentId}): アカウント情報なし`;
  }

  // 氏名整合性チェック
  if (account.name && account.name !== student.studentName) {
    return `生徒 氏名不一致: 履修「${student.studentName}」vs マッピング「${account.name}」(${student.studentId})`;
  }

  // アカウント有効性チェック
  if (account.isActive === false || account.isActive === "FALSE") {
    return `生徒「${student.studentName}」(${student.studentId}): アカウント無効`;
  }

  // メールアドレス存在チェック
  if (!account.email) {
    return `生徒「${student.studentName}」(${student.studentId}): メールアドレスなし`;
  }

  // ドメインチェック
  if (!account.email.endsWith(CONFIG.DOMAIN_SETTINGS.STUDENT_DOMAIN)) {
    return `生徒「${student.studentName}」(${student.studentId}): 無効なドメイン (期待: ${CONFIG.DOMAIN_SETTINGS.STUDENT_DOMAIN}, 実際: ${account.email})`;
  }

  return null; // バリデーション成功
}

/**
 * 教員アカウントのバリデーション
 * @param {Object} teacher - 教員情報
 * @returns {string|null} エラーメッセージ、または null（正常）
 */
function validateTeacherAccount(teacher) {
  // メールアドレス存在チェック
  if (!teacher.teacherEmail) {
    return `教員「${teacher.teacherName}」: メールアドレスなし`;
  }

  // ドメインチェック
  if (!teacher.teacherEmail.endsWith(CONFIG.DOMAIN_SETTINGS.TEACHER_DOMAIN)) {
    return `教員「${teacher.teacherName}」: 無効なドメイン (期待: ${CONFIG.DOMAIN_SETTINGS.TEACHER_DOMAIN}, 実際: ${teacher.teacherEmail})`;
  }

  return null; // バリデーション成功
}

/**
 * 保留中の招待を再送信
 * @returns {Object} 実行結果 {successCount, errorCount, resendCount}
 */
function resendPendingInvitations() {
  console.log("=".repeat(60));
  console.log("保留中招待の再送処理");
  console.log("=".repeat(60));

  let successCount = 0;
  let errorCount = 0;
  let resendCount = 0;

  try {
    // DRY_RUN_MODEを確認
    const dryRunMode = isDryRunMode();
    if (dryRunMode) {
      console.log("⚠️ DRY_RUN_MODE: 実際の再送は行いません");
      SpreadsheetApp.getActiveSpreadsheet().toast('DRY_RUNモードで実行中です', '警告', 5);
    } else {
      console.log("✅ 本番モード: 招待を再送信します");
      SpreadsheetApp.getActiveSpreadsheet().toast('招待を再送信中...', '実行中', 3);
    }

    // Active状態のクラスを取得
    const allClasses = getClassesByState(CONFIG.CLASS_STATE.ACTIVE);
    const activeClasses = allClasses.filter(c => c.isActive === true || c.isActive === "TRUE");

    if (activeClasses.length === 0) {
      console.log("再送対象のクラスはありません");
      return { successCount, errorCount, resendCount };
    }

    console.log(`対象クラス数: ${activeClasses.length}`);

    // 各クラスの保留中招待を再送
    for (const classInfo of activeClasses) {
      try {
        const result = resendClassInvitations(classInfo, dryRunMode);
        resendCount += result.resendCount;

        if (result.resendCount > 0) {
          successCount++;
        }

        // API Rate Limit対策
        if (!dryRunMode && result.resendCount > 0) {
          Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
        }
      } catch (error) {
        console.error(`再送エラー: ${classInfo.subjectName}`, error);
        errorCount++;
      }
    }

    console.log("\n" + "=".repeat(60));
    console.log("再送処理完了");
    console.log(`成功: ${successCount}クラス, エラー: ${errorCount}クラス, 再送数: ${resendCount}件`);
    console.log("=".repeat(60));

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `招待再送完了: ${resendCount}件`,
      '完了',
      5
    );

  } catch (error) {
    console.error("招待再送処理エラー:", error);
    throw error;
  }

  return { successCount, errorCount, resendCount };
}

/**
 * 個別クラスの保留中招待を再送
 * @param {Object} classInfo - クラス情報
 * @param {boolean} dryRunMode - DRY_RUNモードかどうか
 * @returns {Object} 実行結果 {resendCount}
 */
function resendClassInvitations(classInfo, dryRunMode) {
  const courseId = classInfo.classroomId;
  const subjectName = classInfo.subjectName;
  let resendCount = 0;

  debugLog(`招待再送開始: ${subjectName} (${courseId})`);

  try {
    // 保留中の招待を取得
    const pendingInvitations = getPendingInvitations(courseId);

    if (pendingInvitations.length === 0) {
      console.log(`✓ ${subjectName}: 再送対象の招待なし`);
      return { resendCount: 0 };
    }

    console.log(`⚠️ ${subjectName}: 保留中の招待 ${pendingInvitations.length}件を再送します`);

    // 各招待を再送（既存の招待を削除して新規作成）
    for (const invitation of pendingInvitations) {
      try {
        if (dryRunMode) {
          console.log(`[DRY-RUN] 再送予定: ${invitation.role} - ${invitation.userId}`);
          resendCount++;
        } else {
          // 既存の招待を削除
          deleteInvitation(invitation.id);
          debugLog(`招待削除: ${invitation.id}`);

          // 新しい招待を作成
          const newInvitation = {
            userId: invitation.userId,
            courseId: courseId,
            role: invitation.role
          };
          Classroom.Invitations.create(newInvitation);
          debugLog(`招待再送成功: ${invitation.role} - ${invitation.userId}`);

          resendCount++;

          // API Rate Limit対策
          Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
        }
      } catch (error) {
        console.error(`招待再送エラー: ${invitation.userId}`, error);
      }
    }

    // ログ記録
    const modeLabel = dryRunMode ? "（DRY-RUN）" : "";
    logToProcessLog({
      subjectName: subjectName,
      classroomId: courseId,
      result: `成功${modeLabel}`,
      message: `招待${resendCount}件を再送しました`,
      executor: dryRunMode ? "Auto System (DRY-RUN)" : "Auto System"
    });

    console.log(`✓ ${subjectName}: ${resendCount}件再送完了`);

  } catch (error) {
    console.error(`招待再送エラー: ${subjectName}`, error);

    logToProcessLog({
      subjectName: subjectName,
      classroomId: courseId,
      result: "失敗",
      message: `招待再送エラー: ${error.message}`,
      executor: "Auto System"
    });

    throw error;
  }

  return { resendCount };
}

/**
 * 招待を削除
 * @param {string} invitationId - 招待ID
 */
function deleteInvitation(invitationId) {
  return executeWithRetry(() => {
    return Classroom.Invitations.delete(invitationId);
  });
}

/**
 * 保留中の招待を取得（Phase 5と共通化のため）
 * @param {string} courseId - クラスID
 * @returns {Array<Object>} 保留中の招待の配列
 */
function getPendingInvitations(courseId) {
  return executeWithRetry(() => {
    try {
      const response = Classroom.Invitations.list({ courseId: courseId });
      return response.invitations || [];
    } catch (error) {
      if (error.message && error.message.includes("not found")) {
        return [];
      }
      throw error;
    }
  });
}
