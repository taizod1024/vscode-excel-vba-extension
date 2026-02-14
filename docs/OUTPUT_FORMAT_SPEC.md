# 出力フォーマット仕様書

## 概要

Excel VBA Extension のすべてのコマンドで使用される、統一された出力フォーマット仕様です。

## 出力チャネル

### チャネル管理

- **チャネル名**: "Excel VBA"
- **表示タイミング**: コマンド実行開始時に自動表示（フォーカス有）
- **表示方法**: `channel.show(false)` - ユーザーのエディタフォーカスは奪わない

### ライフサイクル

1. コマンド実行
2. チャネルを自動表示（フォーカスなし）
3. コマンド処理
4. 結果をチャネルに出力
5. チャネルは開いたまま（ユーザーが手動で閉じるまで）

## メッセージフォーマット

### 1. コマンド開始メッセージ

```
## [コマンド名]
```

**例**:

```
## Load VBA from Excel Book
```

### 2. 処理詳細（PowerShell出力）

```
- [キー]: [値]
```

**特徴**:

- PowerShell からの出力を `Output:` キーで表示
- 末尾の空白（改行）は自動削除
- 内容例: ファイルパス、処理件数、詳細メッセージ

**例**:

```
- Output: VBA extraction completed - 3 modules processed
- Output: Files saved to: c:\Users\...\book_xlsx\bas
```

### 3. 成功メッセージ

```
[SUCCESS] [アクション説明] ([コンテキスト詳細])
```

**仕様**:

- すべてのコマンドで統一形式
- アクション説明: 実行した処理を簡潔に記述
- コンテキスト詳細: ファイル名、件数、パスなど具体的情報

**例一覧**:

| コマンド                     | メッセージ                                                      |
| ---------------------------- | --------------------------------------------------------------- |
| New Excel Book               | `[SUCCESS] New workbook created (filename.xlsx)`                |
| New Excel Book with CustomUI | `[SUCCESS] New macro workbook created (filename.xlsm)`          |
| Open Excel Book              | `[SUCCESS] Opened in Excel`                                     |
| Load VBA                     | `[SUCCESS] VBA extracted (3 file(s)) to folder`                 |
| Save VBA                     | `[SUCCESS] VBA saved to Excel book`                             |
| Compare VBA                  | `[SUCCESS] No differences found (disk and workbook match)`      |
| Run VBA Sub                  | `[SUCCESS] Sub executed (SubName)`                              |
| Load CSV                     | `[SUCCESS] CSV extracted (folder/csv folder)`                   |
| Save CSV                     | `[SUCCESS] Sheets saved from CSV file`                          |
| Create URL Shortcut          | `[SUCCESS] URL shortcuts created (cloud file references saved)` |
| Load CustomUI                | `[SUCCESS] CustomUI extracted (2 file(s)) to folder`            |
| Save CustomUI                | `[SUCCESS] CustomUI saved to Excel book`                        |
| Export Sheets as PNG         | `[SUCCESS] Sheets exported (PNG images created)`                |
| Open Sheet from PNG          | `[SUCCESS] Opened in Excel`                                     |

### 4. エラーメッセージ

```
[ERROR] [エラー内容]
```

**特徴**:

- エラーは出力チャネルに記録
- 同時にダイアログボックスでユーザーに通知
- 詳細な技術情報も出力チャネルに記録

**例**:

```
[ERROR] No book open. Please open an Excel file in Excel.
[ERROR] Required folder not found: c:\path\to\book_xlsx\bas
```

## 出力例

### Load VBA from Excel Book の完全な出力例

```
2026-02-14 20:24:29.856 [info] ## Load VBA from Excel Book
2026-02-14 20:24:29.856 [info] - Output: VBA modules extracted successfully
[SUCCESS] VBA extracted (3 file(s)) to folder
```

### Create URL Shortcut の完全な出力例

```
2026-02-14 20:24:29.856 [info] ## Create URL Shortcut for Cloud Files
2026-02-14 20:24:29.856 [info] - Output: URL shortcuts created successfully
[SUCCESS] URL shortcuts created (cloud file references saved)
```

### エラーハンドリングの出力例

```
2026-02-14 20:24:29.856 [info] ## Save VBA to Excel Book
2026-02-14 20:24:29.856 [error] [ERROR] No book open. Please open an Excel file in Excel.
```

## ファイルエンコーディング

### VBA ファイル

- **エンコーディング**: Shift_JIS (Shift-JIS)
- **理由**: Excel VBA ネイティブの文字エンコーディング
- **対象**: `.bas`, `.cls`, `.frm` ファイル

### CSV ファイル

- **エンコーディング**: UTF-8
- **BOM**: なし（Excel との互換性）
- **行末**: CRLF（Windows 標準）

### CustomUI XML ファイル

- **エンコーディング**: UTF-8
- **BOM**: なし
- **宣言**: XML 宣言で `utf-8` を指定

### URL ショートカット（.url）

- **エンコーディング**: Shift-JIS（システムデフォルト）
- **形式**: Windows Internet Shortcut
- **コード例**: PowerShell の `Out-File -Encoding Default` で出力

```powershell
$content | Out-File -LiteralPath $shortcutPath -Encoding Default -Force
```

## カラー・アイコン

### メッセージレベル表示

| レベル  | 表示形式    | 用途                       |
| ------- | ----------- | -------------------------- |
| SUCCESS | `[SUCCESS]` | コマンド正常完了           |
| ERROR   | `[ERROR]`   | コマンド失敗               |
| INFO    | `[info]`    | タイムスタンプ付き情報出力 |

### VS Code 出力パネルでの色分け

- `[SUCCESS]` - グリーン（デフォールト）
- `[ERROR]` - レッド（デフォールト）
- `[info]` - グレー（デフォールト）

## パフォーマンス ガイドライン

### メモリ効率

- チャネルは常時オープン状態を維持
- ユーザーが手動で「出力パネルをクリア」するまで、全メッセージを保持
- 大規模な操作（1000+ ファイル）でもメモリ効率を維持

### メッセージ表示速度

- 同期メッセージ表示: < 100ms
- PowerShell 出力の取得と表示: < 500ms（ネットワーク遅延なし）

## 実装上の注意点

### PowerShell 出力の処理

```typescript
// ✅ 正しい実装
if (result.stdout) {
  logger.logDetail("Output", result.stdout.trim()); // 末尾の空白を削除
}

// ❌ 非推奨（空行が生成される）
if (result.stdout) {
  logger.logDetail("Output", result.stdout); // 末尾の改行が残る
}
```

### Logger クラスの使用方法

```typescript
const logger = new Logger(context.channel);

// コマンド開始時
logger.logCommandStart("Command Name");

// 詳細情報
logger.logDetail("Info Key", "value");

// PowerShell 出力（既にtrimされているはず）
if (result.stdout) logger.logDetail("Output", result.stdout.trim());

// 成功メッセージ（必ずlogSuccess()を使用）
logger.logSuccess("Action description (context detail)");

// エラーメッセージ（catch ブロック内）
logger.logError(reason);
```

### チャネル表示のベストプラクティス

```typescript
// コマンド実行時の処理フロー
vscode.commands.registerCommand('command-id', async (...) => {
  const commandContext = { channel, extensionPath };

  // ✅ チャネルを最初に表示
  this.channel.show(false);  // ← ここで表示

  try {
    // コマンド処理
  } catch (reason) {
    // エラーメッセージのみ出力
    // show(false) は不要（既に表示されている）
  }
});
```

## 関連ドキュメント

- [ALL_FEATURES_OVERVIEW.md](./ALL_FEATURES_OVERVIEW.md) - 全機能概要
- [COMMAND_REFERENCE.md](./COMMAND_REFERENCE.md) - コマンドリファレンス
