# Google Classroom 年度更新自動化システム

Google Classroomのクラス管理作業を自動化するシステム（Phase 1実装済み）

## プロジェクト概要

このシステムは、Google Classroomの年度更新作業を自動化し、運用効率の向上とヒューマンエラーの削減を実現します。

### 実装状況

- ✅ **Phase 1: 旧年度クラスのアーカイブ** - 実装済み
- ⏳ Phase 2: 新年度クラスの作成 - 未実装
- ⏳ Phase 3: トピックの作成 - 未実装
- ⏳ Phase 4: 生徒・教員の一括登録 - 未実装

## プロジェクト構成

```
classroom-automation/
├── appsscript.json          # GASマニフェスト
├── src/
│   ├── main.gs              # メインエントリーポイント
│   ├── config.gs            # 設定管理モジュール
│   ├── phases/
│   │   └── phase1_archive.gs      # Phase 1: アーカイブ
│   ├── services/
│   │   └── sheet_service.gs       # スプレッドシート操作
│   └── utils/
│       ├── logger.gs              # ログ記録
│       └── error_handler.gs       # エラー処理
└── .clasp.json              # clasp設定
```

## セットアップ

### 1. スプレッドシート設定

対象スプレッドシート:
- スプレッドシートID: `1BJ-KBSQyJqbb9F8ixBWfXKLlPP4iALfsfuv9DJ-2or4`

#### 必須: システム設定シートに以下を追加

「システム設定 」シート（末尾にスペースあり）に以下の設定を追加してください:

| 設定項目名 | 設定値 | 備考 |
|-----------|--------|------|
| DRY_RUN_MODE | TRUE | テスト実行モード（実際のAPI呼び出しなし） |
| 管理者アカウントID | admin.it@school.jp | クラス作成時のオーナーアカウント（Phase 2で使用） |
| デバッグモード | FALSE | 詳細ログ出力の有効化 |

**重要**: 初回テストでは必ず `DRY_RUN_MODE = TRUE` に設定してください！

### 2. GASプロジェクト

GASプロジェクトID: `1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ`

コードは既にデプロイ済みです。GASエディタで確認できます:
https://script.google.com/d/1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ/edit

### 3. 初回認証

GASエディタで初めて実行する際、以下の権限が要求されます:

- Google Classroom へのアクセス
- Google Sheets へのアクセス
- 外部リクエストの送信

「権限を確認」→「許可」をクリックして承認してください。

## 使用方法

### GASエディタでの実行

1. GASエディタを開く: https://script.google.com/d/1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ/edit

2. 以下の関数を実行できます:

#### テスト用関数

| 関数名 | 説明 |
|--------|------|
| `testSystemSettings()` | システム設定の確認 |
| `testSpreadsheetConnection()` | スプレッドシート接続テスト |
| `testClassMasterData()` | クラスマスタデータの表示 |

#### Phase 1実行関数

| 関数名 | 説明 |
|--------|------|
| `runDryRun()` | DRY-RUNモードで実行（推奨） |
| `testPhase1()` | Phase 1のみテスト実行 |
| `runPhase1Archive()` | Phase 1を実際に実行 |

### 実行手順

#### 1. 動作確認（推奨）

```
1. GASエディタで「testSystemSettings()」を実行
   → DRY_RUN_MODE が TRUE であることを確認

2. 「testSpreadsheetConnection()」を実行
   → すべてのシートが正しく認識されているか確認

3. 「testClassMasterData()」を実行
   → アーカイブ対象のクラスが表示されることを確認
```

#### 2. DRY-RUN実行

```
1. 「runDryRun()」を実行
   → ログに「[DRY-RUN]」が表示され、実際のAPI呼び出しは行われません

2. 「登録処理ログ」シートを確認
   → 処理結果が「成功（DRY-RUN）」として記録されていることを確認
```

#### 3. 本番実行（慎重に！）

```
1. スプレッドシートの「システム設定 」シートで
   DRY_RUN_MODE を FALSE に変更

2. 「testPhase1()」または「runPhase1Archive()」を実行
   → 実際にClassroom APIが呼び出され、クラスがアーカイブされます

3. 「登録処理ログ」シートと「エラー未処理リスト」シートを確認
```

### Phase 1の動作

Phase 1は以下の処理を実行します:

1. クラスマスタから状態が "Archived" のクラスを抽出
2. 各クラスに対して:
   - クラス名に年度プレフィックスを付与（例: "2025_数学"）
   - Classroom APIでクラスをアーカイブ
3. 処理結果を「登録処理ログ」に記録
4. エラーが発生した場合は「エラー未処理リスト」に記録

## ローカル開発

### clasp使用

```bash
# プロジェクトのクローン
clasp clone 1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ

# コードの編集後、GASにプッシュ
clasp push -f

# ログの確認
clasp logs
```

## トラブルシューティング

### DRY_RUN_MODE設定が見つからない

**エラー**: `設定値取得エラー: DRY_RUN_MODE`

**解決策**: スプレッドシートの「システム設定 」シート（末尾にスペースあり）に以下を追加:
- A列: `DRY_RUN_MODE`
- B列: `TRUE`

### シートが見つからない

**エラー**: `シート「XXX」が見つかりません`

**解決策**: 要件定義書に従って、必要なシートを作成してください。必須シート:
- クラス（科目）マスタ
- 履修登録マスタ
- 教員マスタ
- アカウントマッピング
- 登録処理ログ
- エラー未処理リスト
- システム設定 （末尾にスペースあり）
- 一時作業スペース

### 権限エラー

**エラー**: `権限不足`

**解決策**: GASエディタで関数を実行し、権限の許可を求められたら承認してください。

## 次のステップ

Phase 1が正常に動作することを確認したら、以下の実装を進めます:

1. Phase 2: 新年度クラスの作成
2. Phase 3: トピックの作成
3. Phase 4: 生徒・教員の一括登録

## 参考資料

- [要件定義書](./Google_Classroom_年度更新自動化システム_要件定義書_v2.md)
- [Google Classroom API](https://developers.google.com/classroom)
- [clasp](https://github.com/google/clasp)
