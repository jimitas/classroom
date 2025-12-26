# 開発ログ - Google Classroom 年度更新自動化システム

## プロジェクト情報

- **開始日**: 2025年12月26日
- **GASプロジェクトID**: `1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ`
- **スプレッドシートID**: `1BJ-KBSQyJqbb9F8ixBWfXKLlPP4iALfsfuv9DJ-2or4`
- **GitHubリポジトリ**: https://github.com/jimitas/classroom.git

---

## 実装状況

### Phase 1: 旧年度クラスのアーカイブ ✅ 完了

**実装日**: 2025年12月26日

#### 機能概要
- クラスマスタから状態が "Archived" のクラスを自動抽出
- クラス名に年度プレフィックスを付与（例: "テスト２" → "2025_テスト２"）
- Classroom APIでクラスをアーカイブ
- DRY_RUNモードでテスト実行可能
- エラーハンドリングとログ記録

#### 実装ファイル
- `src/config.gs` - 設定管理
- `src/main.gs` - メインエントリーポイント
- `src/phases/phase1_archive.gs` - Phase 1実装
- `src/services/sheet_service.gs` - スプレッドシート操作
- `src/utils/logger.gs` - ログ記録
- `src/utils/error_handler.gs` - エラー処理
- `appsscript.json` - GASマニフェスト

---

### ユーティリティ機能: クラスマスタ同期 ✅ 完了

**実装日**: 2025年12月26日

#### 機能概要
- Google Classroomの現在の状態をクラスマスタシートに自動同期
- 既存クラスの更新と新規クラスの追加に対応
- クラス状態（Active/Archived）を自動判定
- 既存の手動入力データ（履修登録マスタの列名、担当教員名など）を保持

#### 主要関数
- `syncClassMasterFromClassroom()` - Classroomの全クラスをクラスマスタに同期
- `listAllMyCourses()` - 自分が教師として参加している全クラスを一覧表示
- `getCourseIdFromEnrollmentCode()` - 招待コードからクラスIDを取得
- `mapCourseState()` - Classroom APIのcourseStateをマッピング

#### クラス状態マッピング

| Classroom API | クラスマスタ | 説明 |
|--------------|------------|------|
| ACTIVE | Active | 運用中 |
| ARCHIVED | Archived | アーカイブ済み |
| PROVISIONED | New | 作成されたが未公開 |
| DECLINED | Archived | 招待が拒否された |
| SUSPENDED | Archived | 一時停止 |

#### 科目有効性について

**D列「科目有効性」の役割**:
- **Phase 2（クラス作成）**: `TRUE` の科目のみクラスを作成
- **Phase 4（履修登録）**: `TRUE` の科目のみ履修登録処理を実行
- **用途例**: 休講科目や廃止科目を `FALSE` に設定

---

## セットアップ手順

### 1. clasp設定

```bash
# 既存のGASプロジェクトをクローン
clasp clone 1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ

# コード編集後、GASにプッシュ
clasp push -f
```

### 2. スプレッドシート設定

「システム設定 」シート（末尾にスペースあり）に以下を追加:

| 設定項目名 | 設定値 | 備考 |
|-----------|--------|------|
| DRY_RUN_MODE | TRUE | テスト実行モード（実際のAPI呼び出しなし） |
| 管理者アカウントID | admin.it@school.jp | クラス作成時のオーナーアカウント（Phase 2で使用） |
| デバッグモード | FALSE | 詳細ログ出力の有効化 |

### 3. GASエディタでの初回認証

GASエディタ: https://script.google.com/d/1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ/edit

初回実行時に以下の権限を許可:
- Google Classroom へのアクセス
- Google Sheets へのアクセス
- 外部リクエストの送信

---

## 遭遇した問題と解決方法

### 問題1: 招待コードとクラスIDの混同

**症状**:
- URL `https://classroom.google.com/r/ODIzOTk5OTI1Mzcy` の `ODIzOTk5OTI1Mzcy` をクラスIDと誤認識

**原因**:
- `/r/` で始まるURLは招待コード（Enrollment Code）
- クラスID（Course ID）とは異なる

**解決策**:
- `listAllMyCourses()` 関数を追加
- `getCourseIdFromEnrollmentCode()` 関数を追加
- これらの関数で正しいクラスIDを取得

**実装コミット**: `e748ea7`

---

### 問題2: Classroom API `courses.update` のエラー

**症状**:
```
API call to classroom.courses.update failed with error:
course.name: Course name must be specified.
```

**原因**:
- Google Classroom APIの`courses.update`は部分更新をサポートしていない
- クラス名を指定せずに`courseState`だけを更新しようとした

**元のコード**:
```javascript
// クラス名を変更
updateClassroomName(classroomId, archivedName);

// クラスをアーカイブ（エラー発生！）
archiveClassroom(classroomId) {
  const course = {
    courseState: "ARCHIVED"  // nameが無い！
  };
  return Classroom.Courses.update(course, courseId);
}
```

**解決策**:
- 2回のAPI呼び出しを1回に統合
- クラス名とアーカイブ状態を同時に指定

**修正後のコード**:
```javascript
function updateAndArchiveClassroom(courseId, newName) {
  const course = {
    name: newName,
    courseState: "ARCHIVED"
  };
  return Classroom.Courses.update(course, courseId);
}
```

**実装コミット**: `a586773`

---

### 問題3: `course.id` フィールドのエラー

**症状**:
```
Invalid value at 'course.id' (TYPE_STRING), 823999925372
```

**原因**:
- `Classroom.Courses.update(resource, id)` の呼び出しで、`resource`オブジェクトに`id`フィールドを含めていた
- Google Classroom APIでは、リクエストボディに`id`を含めてはいけない

**元のコード**:
```javascript
const course = {
  id: courseId,        // これが原因！
  name: newName,
  courseState: "ARCHIVED"
};
Classroom.Courses.update(course, courseId);
```

**解決策**:
- `resource`オブジェクトから`id`フィールドを削除
- `id`は第2引数として渡すだけでOK

**修正後のコード**:
```javascript
const course = {
  name: newName,       // idフィールドを削除
  courseState: "ARCHIVED"
};
Classroom.Courses.update(course, courseId);
```

**実装コミット**: `6458427`

---

## テスト結果

### Phase 1アーカイブ機能

#### テスト環境
- テストクラス: "テスト２"
- クラスID: `823999925372`
- 招待コード: `ODIzOTk5OTI1Mzcy`

#### テスト手順
1. `testSystemSettings()` - システム設定確認 ✅
2. `testSpreadsheetConnection()` - スプレッドシート接続確認 ✅
3. `testClassMasterData()` - クラスマスタデータ確認 ✅
4. `listAllMyCourses()` - クラス一覧表示とクラスID取得 ✅
5. クラスマスタに「テスト２」を追加（状態: Archived）
6. `runDryRun()` - DRY-RUN実行 ✅
7. `DRY_RUN_MODE` を `FALSE` に変更
8. `testPhase1()` - 本番実行 ✅

#### 結果
- ✅ クラス名が "テスト２" → "2025_テスト２" に変更
- ✅ クラスがアーカイブ状態に変更
- ✅ 登録処理ログに成功記録
- ✅ エラーなし

---

## 主要な関数一覧

### メインエントリーポイント（`src/main.gs`）

| 関数名 | 説明 |
|--------|------|
| `runAllPhases()` | 全フェーズを順次実行 |
| `testPhase1()` | Phase 1のみテスト実行 |
| `runDryRun()` | DRY-RUNモードで実行 |
| `testSystemSettings()` | システム設定の確認 |
| `testSpreadsheetConnection()` | スプレッドシート接続テスト |
| `testClassMasterData()` | クラスマスタデータの表示 |
| `listAllMyCourses()` | 自分が教師のクラス一覧を表示 |
| `getCourseIdFromEnrollmentCode(code)` | 招待コードからクラスIDを取得 |
| `syncClassMasterFromClassroom()` | **NEW** Classroomの状態をクラスマスタに同期 |
| `mapCourseState(apiState)` | **NEW** Classroom APIのcourseStateをマッピング |

### Phase 1（`src/phases/phase1_archive.gs`）

| 関数名 | 説明 |
|--------|------|
| `runPhase1Archive()` | Phase 1メイン処理 |
| `archiveClass(classInfo, dryRunMode)` | 個別クラスのアーカイブ |
| `updateAndArchiveClassroom(courseId, newName)` | クラス名変更とアーカイブ |

### 設定管理（`src/config.gs`）

| 関数名 | 説明 |
|--------|------|
| `getSystemSetting(settingName)` | システム設定値を取得 |
| `isDryRunMode()` | DRY_RUN_MODEかどうか確認 |
| `isDebugMode()` | デバッグモードかどうか確認 |
| `getAdminAccountId()` | 管理者アカウントIDを取得 |

### ログ記録（`src/utils/logger.gs`）

| 関数名 | 説明 |
|--------|------|
| `logToProcessLog(logData)` | 登録処理ログに記録 |
| `logToErrorList(errorData)` | エラー未処理リストに記録 |
| `debugLog(message, data)` | デバッグログ出力 |
| `logPhaseStart(phaseName)` | フェーズ開始ログ |
| `logPhaseComplete(phaseName, success, error)` | フェーズ完了ログ |

### エラー処理（`src/utils/error_handler.gs`）

| 関数名 | 説明 |
|--------|------|
| `executeWithRetry(apiFunction, maxRetries)` | リトライ付きAPI実行 |
| `isRateLimitError(error)` | Rate Limitエラーか判定 |
| `isRetryableError(error)` | リトライ可能エラーか判定 |
| `getErrorCode(error)` | エラーコードを特定 |
| `handleError(error, context)` | エラーを処理してログ記録 |

### スプレッドシート操作（`src/services/sheet_service.gs`）

| 関数名 | 説明 |
|--------|------|
| `getSheetData(sheetName)` | シートデータを取得 |
| `appendToSheet(sheetName, rowData)` | シートに行を追加 |
| `getClassesByState(classState)` | 特定状態のクラスを取得 |
| `updateClassMaster(rowIndex, updates)` | クラスマスタを更新 |
| `getEnrollmentData()` | 履修登録データ取得（縦持ち変換） |
| `getTeacherData()` | 教員担当データ取得（縦持ち変換） |
| `getStudentAccount(studentId)` | 学籍番号からGoogleアカウント取得 |

---

## Google Classroom APIの重要な仕様

### `Classroom.Courses.update()` の正しい使い方

#### ❌ 間違った使い方

```javascript
// 部分更新はできない
const course = {
  courseState: "ARCHIVED"  // nameが無いとエラー！
};
Classroom.Courses.update(course, courseId);
```

```javascript
// idフィールドを含めるとエラー
const course = {
  id: courseId,  // これを含めてはいけない！
  name: "新しい名前",
  courseState: "ARCHIVED"
};
Classroom.Courses.update(course, courseId);
```

#### ✅ 正しい使い方

```javascript
// 必要なフィールドをすべて指定
const course = {
  name: "新しいクラス名",  // 必須
  courseState: "ARCHIVED"
};
Classroom.Courses.update(course, courseId);
```

---

## 今後の開発予定

### Phase 2: 新年度クラスの作成

**実装予定内容**:
- クラス状態が "New" のクラスを抽出
- 管理者アカウントをオーナーとしてクラス作成
- 作成後、クラスIDをクラスマスタに記録
- クラス状態を "Active" に更新

**必要な実装**:
- `src/phases/phase2_create.gs` の作成
- `Classroom.Courses.create()` API呼び出し
- クラスマスタの更新処理

---

### Phase 3: トピックの作成

**実装予定内容**:
- クラス状態が "Active" のクラスを対象
- 固定トピックパターンを逆順で作成:
  1. お知らせ
  2. 授業資料
  3. 課題
  4. テスト
  5. その他

**必要な実装**:
- `src/phases/phase3_topics.gs` の作成
- `Classroom.Courses.Topics.create()` API呼び出し

---

### Phase 4: 生徒・教員の一括登録

**実装予定内容**:
- 履修登録マスタから生徒の履修情報を取得（マトリクス→縦持ち変換）
- 教員マスタから教員の担当情報を取得（マトリクス→縦持ち変換）
- アカウントマッピングからGoogleアカウント情報を取得
- バッチ処理でClassroom APIに登録

**必要な実装**:
- `src/phases/phase4_members.gs` の作成
- `src/services/data_transformer.gs` の作成（マトリクス変換）
- `Classroom.Courses.Students.create()` API呼び出し
- `Classroom.Courses.Teachers.create()` API呼び出し

---

## 開発のベストプラクティス

### 1. DRY_RUNモードを活用

本番実行前に必ず`DRY_RUN_MODE = TRUE`でテスト実行し、ログを確認する。

### 2. エラーハンドリング

すべてのAPI呼び出しは`executeWithRetry()`でラップし、Rate Limit対策を行う。

### 3. ログ記録

処理結果は必ず「登録処理ログ」に記録し、エラーは「エラー未処理リスト」に記録する。

### 4. API Rate Limit対策

連続したAPI呼び出しの間には`Utilities.sleep(300)`で待機時間を挿入する。

### 5. Git管理

機能追加・バグ修正ごとに適切なコミットメッセージでコミットする。

---

## デバッグのヒント

### GASエディタでのログ確認

1. GASエディタで関数を実行
2. 下部の「実行ログ」タブを確認
3. または `clasp logs` コマンドでローカルから確認

### スプレッドシートでの確認

- 登録処理ログシート: すべての処理結果を確認
- エラー未処理リストシート: エラー発生時の詳細を確認

### デバッグモードの有効化

システム設定シートで「デバッグモード」を`TRUE`に設定すると、詳細ログが出力される。

---

## 参考リンク

- [Google Classroom API Reference](https://developers.google.com/classroom/reference/rest)
- [clasp GitHub](https://github.com/google/clasp)
- [Google Apps Script Reference](https://developers.google.com/apps-script/reference)
- [要件定義書](./Google_Classroom_年度更新自動化システム_要件定義書_v2.md)

---

## コミット履歴

| コミットID | 日付 | 内容 |
|-----------|------|------|
| `7f1e030` | 2025-12-26 | Initial commit: 要件定義書とプロジェクト設定を追加 |
| `2c4ec47` | 2025-12-26 | Implement Phase 1: Archive functionality |
| `e748ea7` | 2025-12-26 | Add helper functions to get Course ID from enrollment code |
| `a586773` | 2025-12-26 | Fix Phase 1: Combine name update and archive into single API call |
| `6458427` | 2025-12-26 | Fix: Remove id field from courses.update request body |
| `eab1849` | 2025-12-26 | Add comprehensive development log (DEVELOPMENT.md) |
| `9a0a3d4` | 2025-12-26 | Add syncClassMasterFromClassroom() to sync Classroom state to spreadsheet |

---

**最終更新**: 2025年12月26日
**Phase 1実装完了**: ✅
**ユーティリティ機能実装完了**: ✅（クラスマスタ同期）
**次のステップ**: Phase 2の実装
