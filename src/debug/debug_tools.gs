/**
 * デバッグ・診断ツール
 * @module DebugTools
 *
 * このファイルには開発・トラブルシューティング用の関数が含まれています。
 * 本番環境では使用しませんが、問題発生時の診断に役立ちます。
 */

/**
 * クラスのオーナー権限を確認
 * @param {string} courseId - クラスID（省略時はデフォルトID使用）
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
 * Phase 4のデータ変換結果を表示（デバッグ用）
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
