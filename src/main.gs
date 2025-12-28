/**
 * Google Classroom 年度更新自動化システム
 * メインエントリーポイント
 * @module Main
 */

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 Classroom自動化')
    .addSubMenu(ui.createMenu('🧪 テスト実行（DRY-RUN）')
      .addItem('Phase 1: アーカイブテスト', 'testPhase1')
      .addItem('Phase 2: クラス作成テスト', 'testPhase2')
      .addItem('Phase 3: トピック作成テスト', 'testPhase3')
      .addItem('Phase 4: メンバー招待テスト', 'testPhase4'))
    .addSeparator()
    .addSubMenu(ui.createMenu('▶️ 本番実行')
      .addItem('Phase 1: 旧年度クラスをアーカイブ', 'runPhase1Archive')
      .addItem('Phase 2: 新年度クラスを作成', 'runPhase2CreateClasses')
      .addItem('Phase 3: トピックを作成', 'runPhase3CreateTopics')
      .addItem('Phase 4: 生徒・教員を招待', 'runPhase4RegisterMembers'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 ユーティリティ')
      .addItem('Classroomからクラスマスタを同期', 'syncClassMasterFromClassroom')
      .addItem('使い方ガイドを更新', 'updateUsageGuideSheet'))
    .addToUi();
}

/**
 * 使い方ガイドをスプレッドシートに書き込み
 */
function updateUsageGuideSheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('使い方ガイド');

    if (!sheet) {
      throw new Error('シート「使い方ガイド」が見つかりません。先にシートを作成してください。');
    }

    // シートをクリア
    sheet.clear();

    // ヘッダー行の設定
    sheet.getRange('A1:B1').setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12);

    // 使い方ガイドの内容
    const guideData = [
      ['Google Classroom 年度更新自動化システム - 使い方ガイド', ''],
      ['', ''],
      ['📋 基本操作', ''],
      ['メニューからの実行', 'スプレッドシート上部の「📚 Classroom自動化」メニューから各機能を実行できます'],
      ['', ''],
      ['🧪 テスト実行の手順', ''],
      ['1. DRY_RUNモードを有効化', '「システム設定 」シートでDRY_RUN_MODEをTRUEに設定'],
      ['2. テスト実行', 'メニュー → 🧪 テスト実行 → 各Phaseを選択'],
      ['3. 結果確認', '「登録処理ログ」シートで結果を確認（DRY-RUN表示あり）'],
      ['', ''],
      ['▶️ 本番実行の手順', ''],
      ['1. DRY_RUNモードを無効化', '「システム設定 」シートでDRY_RUN_MODEをFALSEに設定'],
      ['2. 本番実行', 'メニュー → ▶️ 本番実行 → 各Phaseを選択'],
      ['3. 結果確認', '「登録処理ログ」と「エラー未処理リスト」シートを確認'],
      ['', ''],
      ['📝 各Phaseの説明', ''],
      ['Phase 1: 旧年度クラスのアーカイブ', 'クラスマスタで状態が「Archived」のクラスを自動アーカイブ'],
      ['', '→ クラス名に年度プレフィックスを付与（例: 2025_数学）'],
      ['Phase 2: 新年度クラスの作成', 'クラスマスタで状態が「New」かつ科目有効性が「TRUE」のクラスを作成'],
      ['', '→ 作成後、クラスIDを記録し状態を「Active」に更新'],
      ['Phase 3: トピックの作成', '状態が「Active」かつトピック作成フラグが「TRUE」のクラスにトピックを作成'],
      ['', '→ 5つの標準トピック（お知らせ、授業資料、課題、テスト、その他）'],
      ['Phase 4: 生徒・教員の招待', '履修登録マスタ・教員マスタから対象者を抽出し、招待メールを送信'],
      ['', '→ 生徒・教員は招待メールの「承諾」ボタンをクリックして参加'],
      ['', '→ 履修登録マスタのC列「有効」フラグがTRUEの生徒のみ招待'],
      ['', ''],
      ['📌 転入生の追加方法（Phase 4）', ''],
      ['1. 履修登録マスタで既登録者のC列を「FALSE」に変更', '既に招待済みの生徒をスキップ対象にする'],
      ['2. 新しい転入生の行を追加', '氏名・学籍番号・有効=TRUE・履修科目=○'],
      ['3. Phase 4を実行', '転入生のみが招待される（既登録者はスキップ）'],
      ['4. 完了後、全員のC列を「TRUE」に戻す', '次回の一括処理に備える'],
      ['', ''],
      ['🔧 ユーティリティ機能', ''],
      ['Classroomからクラスマスタを同期', 'Classroomの現在の状態をスプレッドシートに反映'],
      ['', '→ クラス名、状態（Active/Archived）を自動更新'],
      ['使い方ガイドを更新', 'この使い方ガイドを最新の内容に更新'],
      ['', ''],
      ['⚠️ 重要な注意事項', ''],
      ['DRY_RUNモードの確認', '本番実行前に必ずDRY_RUNモードでテストしてください'],
      ['招待の承諾', 'Phase 4実行後、生徒・教員は招待メールを承諾する必要があります'],
      ['エラーの確認', '実行後は必ず「登録処理ログ」と「エラー未処理リスト」を確認してください'],
      ['', ''],
      ['📊 ログの見方', ''],
      ['登録処理ログ', '全ての処理結果が記録されます（成功・失敗・スキップ）'],
      ['エラー未処理リスト', 'エラーが発生した場合の詳細が記録されます'],
      ['', ''],
      ['🆘 トラブルシューティング', ''],
      ['DRY_RUN_MODE設定が見つからない', '「システム設定 」シートにA列「DRY_RUN_MODE」、B列「TRUE」を追加'],
      ['シートが見つかりません', '要件定義書に従って必須シートを作成してください'],
      ['権限不足エラー', 'GASエディタで関数を実行し、権限の許可を承認してください'],
      ['', ''],
      ['📚 参考資料', ''],
      ['詳細な使い方', 'README.mdを参照'],
      ['開発ログ', 'DEVELOPMENT.mdを参照'],
      ['今後の機能追加予定', 'FUTURE_FEATURES.mdを参照'],
      ['', ''],
      ['最終更新', new Date().toLocaleString('ja-JP')]
    ];

    // データを書き込み
    sheet.getRange(1, 1, guideData.length, 2).setValues(guideData);

    // 列幅を調整
    sheet.setColumnWidth(1, 350);
    sheet.setColumnWidth(2, 600);

    // タイトル行のスタイル
    sheet.getRange('A1:B1').merge().setHorizontalAlignment('center');

    // セクションヘッダーのスタイル（絵文字で始まる行）
    for (let i = 1; i <= guideData.length; i++) {
      const cellValue = sheet.getRange(i, 1).getValue();
      if (cellValue && (cellValue.toString().startsWith('📋') ||
                        cellValue.toString().startsWith('🧪') ||
                        cellValue.toString().startsWith('▶️') ||
                        cellValue.toString().startsWith('📝') ||
                        cellValue.toString().startsWith('🔧') ||
                        cellValue.toString().startsWith('⚠️') ||
                        cellValue.toString().startsWith('📊') ||
                        cellValue.toString().startsWith('🆘') ||
                        cellValue.toString().startsWith('📚'))) {
        sheet.getRange(i, 1, 1, 2).setBackground('#E8F0FE').setFontWeight('bold').setFontSize(11);
      }
    }

    // Phase説明行のインデント
    for (let i = 1; i <= guideData.length; i++) {
      const cellValue = sheet.getRange(i, 2).getValue();
      if (cellValue && cellValue.toString().startsWith('→')) {
        sheet.getRange(i, 2).setFontColor('#666666').setFontStyle('italic');
      }
    }

    // 枠線を追加
    sheet.getRange(1, 1, guideData.length, 2).setBorder(true, true, true, true, true, true);

    console.log('✅ 使い方ガイドを更新しました');
    SpreadsheetApp.getActiveSpreadsheet().toast('使い方ガイドを更新しました', '成功', 3);

  } catch (error) {
    console.error('使い方ガイド更新エラー:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('エラー: ' + error.message, 'エラー', 5);
    throw error;
  }
}

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
 * テスト実行の共通処理
 * @param {number} phaseNumber - フェーズ番号
 * @param {Function} phaseFunction - 実行するフェーズ関数
 */
function executePhaseTest(phaseNumber, phaseFunction) {
  console.log("=".repeat(60));
  console.log(`🧪 Phase ${phaseNumber} テスト実行（強制DRY_RUNモード）`);
  console.log("=".repeat(60));
  console.log("⚠️ スプレッドシートのDRY_RUN_MODE設定に関わらず、");
  console.log("   このテストは実際のAPI呼び出しを行いません。");
  console.log("");

  // DRY_RUNモードを一時的に強制
  setForceDryRunMode(true);

  try {
    const result = phaseFunction();
    console.log("");
    console.log(`✅ Phase ${phaseNumber}テスト完了:`, result);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Phase ${phaseNumber}テスト完了（DRY-RUN）`, '成功', 3);
  } finally {
    // 強制設定を解除
    setForceDryRunMode(false);
  }
}

/**
 * Phase 1のみを実行（テスト用 - 強制DRY_RUNモード）
 */
function testPhase1() {
  executePhaseTest(1, runPhase1Archive);
}

/**
 * Phase 2のみを実行（テスト用 - 強制DRY_RUNモード）
 */
function testPhase2() {
  executePhaseTest(2, runPhase2CreateClasses);
}

/**
 * Phase 3のみを実行（テスト用 - 強制DRY_RUNモード）
 */
function testPhase3() {
  executePhaseTest(3, runPhase3CreateTopics);
}

/**
 * Phase 4のみを実行（テスト用 - 強制DRY_RUNモード）
 */
function testPhase4() {
  executePhaseTest(4, runPhase4RegisterMembers);
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
/**
 * デバッグ関数は src/debug/debug_tools.gs に移動しました
 *
 * 利用可能なデバッグ関数:
 * - checkClassOwner(courseId) - クラスオーナー権限の確認
 * - debugPhase4Data() - Phase 4データ変換結果の表示
 * - diagnoseApiPermission() - API権限の詳細診断
 */
