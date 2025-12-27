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

      // API Rate Limit対策
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

  if (dryRunMode) {
    // DRY_RUNモード: ログのみ
    console.log(`[DRY-RUN] ${subjectName}:`);
    console.log(`  生徒: ${students.length}名, 教員: ${teachers.length}名`);

    logToProcessLog({
      subjectName: subjectName,
      classroomId: classId,
      result: "成功（DRY-RUN）",
      message: `生徒${students.length}名、教員${teachers.length}名を登録予定`,
      executor: "Auto System (DRY-RUN)"
    });

    return { skipped: false, studentCount: students.length, teacherCount: teachers.length };
  }

  // 実際の登録処理
  // 生徒を登録
  students.forEach(student => {
    const account = accountMap[student.studentId];

    if (!account) {
      errors.push(`生徒「${student.studentName}」(${student.studentId}): アカウント情報なし`);
      return;
    }

    if (account.isActive === false || account.isActive === "FALSE") {
      errors.push(`生徒「${student.studentName}」(${student.studentId}): アカウント無効`);
      return;
    }

    if (!account.email) {
      errors.push(`生徒「${student.studentName}」(${student.studentId}): メールアドレスなし`);
      return;
    }

    try {
      addStudentToClass(classId, account.email);
      studentCount++;
      debugLog(`生徒登録成功: ${student.studentName} → ${subjectName}`);

      // API間隔
      Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
    } catch (error) {
      errors.push(`生徒「${student.studentName}」: ${error.message}`);
    }
  });

  // 教員を登録
  teachers.forEach(teacher => {
    if (!teacher.teacherEmail) {
      errors.push(`教員「${teacher.teacherName}」: メールアドレスなし`);
      return;
    }

    try {
      addTeacherToClass(classId, teacher.teacherEmail);
      teacherCount++;
      debugLog(`教員登録成功: ${teacher.teacherName} → ${subjectName}`);

      // API間隔
      Utilities.sleep(CONFIG.API_SETTINGS.REQUEST_INTERVAL_MS);
    } catch (error) {
      errors.push(`教員「${teacher.teacherName}」: ${error.message}`);
    }
  });

  // ログ記録
  const message = errors.length > 0
    ? `生徒${studentCount}名、教員${teacherCount}名を登録。エラー${errors.length}件: ${errors.join('; ')}`
    : `生徒${studentCount}名、教員${teacherCount}名を登録しました`;

  logToProcessLog({
    subjectName: subjectName,
    classroomId: classId,
    result: errors.length > 0 ? "一部失敗" : "成功",
    message: message,
    executor: "Auto System"
  });

  console.log(`✓ ${subjectName}: 生徒${studentCount}名, 教員${teacherCount}名登録`);

  return { skipped: false, studentCount, teacherCount };
}

/**
 * 生徒をクラスに追加
 * @param {string} courseId - クラスID
 * @param {string} studentEmail - 生徒のメールアドレス
 */
function addStudentToClass(courseId, studentEmail) {
  return executeWithRetry(() => {
    const student = {
      userId: studentEmail
    };
    return Classroom.Courses.Students.create(student, courseId);
  });
}

/**
 * 教員をクラスに追加（共同教師として）
 * @param {string} courseId - クラスID
 * @param {string} teacherEmail - 教員のメールアドレス
 */
function addTeacherToClass(courseId, teacherEmail) {
  return executeWithRetry(() => {
    const teacher = {
      userId: teacherEmail
    };
    return Classroom.Courses.Teachers.create(teacher, courseId);
  });
}
