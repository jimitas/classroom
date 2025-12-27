# Google Workspace API アクセス設定依頼書

**依頼日**: 2025年12月27日
**依頼者**: 野中孝利 (t-nonaka@office.kyoto-art.ac.jp)
**依頼先**: DX部門（Google Workspace管理者）

---

## 依頼内容の概要

Google Classroom 年度更新自動化システム（Apps Script）において、Google Classroom API経由で生徒・教員をクラスに追加する際に、以下のエラーが発生しています。

```
HTTP 403: The caller does not have permission
```

OAuth認証スコープは正しく設定されており、スクリプト実行者はクラスのオーナーであることも確認済みですが、API経由での生徒・教員の追加が「権限不足」でブロックされています。

Google Workspace管理コンソールでの設定確認・変更をお願いします。

---

## 技術的背景

### 現在の状況

| 項目 | 状態 |
|------|------|
| OAuth Scopes | ✅ `classroom.rosters` を含む全スコープ設定済み |
| スクリプト実行者 | ✅ クラスオーナー権限あり |
| クラス状態 | ✅ ACTIVE |
| 手動操作 | ✅ Classroom UIから手動招待は可能 |
| API操作 | ❌ `Students.create()` で HTTP 403 エラー |

### 使用しているAPI

- **Google Classroom API v1**
  - `Classroom.Courses.Students.create()` - 生徒の直接追加
  - `Classroom.Courses.Teachers.create()` - 教員の追加

### エラー詳細

```json
{
  "name": "GoogleJsonResponseException",
  "details": {
    "message": "The caller does not have permission",
    "code": 403
  }
}
```

---

## 確認・変更をお願いしたい設定

Google Workspace管理コンソール (https://admin.google.com) で以下の設定を確認し、必要に応じて変更をお願いします。

### 1. Classroom API のアクセス許可

**パス**: `セキュリティ` → `アクセスとデータ管理` → `APIの制御`

**確認事項**:
- [ ] **Google Classroom API** が組織全体で有効になっているか
- [ ] **内部アプリ（Apps Script）** のAPIアクセスが許可されているか
- [ ] **信頼できる内部アプリ** として登録されているか

### 2. Classroom アプリの設定

**パス**: `アプリ` → `Google Workspace` → `Google Classroom` → `詳細設定`

**確認事項**:
- [ ] **API アクセス** が有効になっているか
- [ ] **クラスの名簿管理** が許可されているか

### 3. Apps Script のアクセス設定

**パス**: `セキュリティ` → `アクセスとデータ管理` → `APIの制御` → `アプリアクセス制御`

**確認事項**:
- [ ] **Apps Script** が信頼できるアプリとして登録されているか
- [ ] スクリプトID `1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ` のアクセスが許可されているか

### 4. ドメイン間共有設定

**パス**: `アプリ` → `Google Workspace` → `Google Classroom` → `共有設定`

**確認事項**:
- [ ] **異なる組織単位間での名簿管理** が許可されているか
  - 教員ドメイン: `@office.kyoto-art.ac.jp`
  - 生徒ドメイン: `@shs.kyoto-art.ac.jp`

---

## 代替案（もし上記設定変更が難しい場合）

上記の設定変更が組織のセキュリティポリシー上難しい場合、以下の代替案を実装予定です:

### Invitations API への切り替え

- `Students.create()` （直接追加）→ `Invitations.create()` （招待メール送信）
- 生徒が招待メールを受信し、「承諾」ボタンをクリックして参加
- 手動招待と同じフローのため、より緩い権限で動作
- **デメリット**: 生徒の承諾アクションが必要

---

## 補足情報

### システム情報
- **プロジェクト名**: Google Classroom 年度更新自動化システム
- **GASプロジェクトID**: `1lFH4ilIL4nczSNGaYikqvh3oQ1FO5J4moIAu367vxnhIzTKmgPxyhoSQ`
- **スプレッドシートID**: `1BJ-KBSQyJqbb9F8ixBWfXKLlPP4iALfsfuv9DJ-2or4`

### OAuth Scopes（設定済み）
```json
[
  "https://www.googleapis.com/auth/classroom.courses",
  "https://www.googleapis.com/auth/classroom.topics",
  "https://www.googleapis.com/auth/classroom.rosters",
  "https://www.googleapis.com/auth/classroom.profile.emails",
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/script.external_request"
]
```

---

## 参考資料

- [Google Classroom API - Students: create](https://developers.google.com/classroom/reference/rest/v1/courses.students/create)
- [Google Workspace Admin - API Controls](https://support.google.com/a/answer/7281227)
- [Google Classroom API - Authorization](https://developers.google.com/classroom/guides/auth)

---

## お問い合わせ

この依頼に関するご質問は、以下までご連絡ください:

- **担当者**: 野中孝利
- **メール**: t-nonaka@office.kyoto-art.ac.jp

ご対応のほど、よろしくお願いいたします。
