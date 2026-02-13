# New Excel Book 機能仕様書

## 概要

VS Code から新しい Excel ファイルを作成する機能です。

## 機能一覧

1. **New Excel Book** - 標準的な Excel ワークブック（.xlsx）を作成
2. **New Excel Book with CustomUI** - CustomUI テンプレート付き Excel ファイル（.xlsm）を作成

## 1. New Excel Book

### 概要

マクロなしの基本的な Excel ファイル（.xlsx）を作成します。

### 入力仕様

**ファイル名入力**
- 拡張子なしの名前を入力
- 日本語を含む各国語対応
- 例: `MyBook`, `売上報告`, `Dashboard`

### 出力仕様

**出力ファイル**
```
{入力名}.xlsx
```

### 処理フロー

1. VS Code のメニューから「File → New File」を選択
2. 表示されるオプションから「New Excel Book」を選択
3. ファイル名を入力
4. Excel アプリケーションが起動
5. 新しいブックが開かれる
6. ワークスペースに保存

### 実装詳細

**メイン処理**: `src/commands/newBook.ts`

```typescript
export async function newBookAsync(context: CommandContext)
```

**PowerShell**: `bin/New-Excel.ps1`

処理内容：
- Excel アプリケーションの取得（存在しない場合は新規作成）
- 新規ブック作成
- .xlsx 形式で保存
- ブックの表示

## 2. New Excel Book with CustomUI

### 概要

CustomUI（リボンカスタマイズ）のテンプレート付き Excel マクロ有効ブック（.xlsm）を作成します。

### 入力仕様

**ファイル名入力**
- 拡張子なしの名前を入力
- 日本語を含む各国語対応
- 例: `MyMacro`, `カスタムUI`, `Dashboard`

### 出力仕様

**出力ファイル**
```
{入力名}.xlsm
  ├── VBA (空のプロジェクト)
  └── CustomUI テンプレート
      ├── customUI.xml (Office 2010+)
      └── customUI14.xml (Office 2010+)
```

### テンプレート内容

**customUI.xml**
基本的なリボンのカスタマイズテンプレート

**customUI14.xml**
Office 2010 以降の新しいリボン定義

### 処理フロー

1. コマンドパレット（Ctrl+Shift+P）を開く
2. 「New Excel Book with CustomUI」を検索して実行
3. ファイル名を入力
4. Excel テンプレートファイル（`excel/addin/excel-vba-addin.xlam`）をコピー
5. ファイルを .xlsm に変換
6. 新規ブックとして Excel で開く

### 実装詳細

**メイン処理**: `src/commands/newBookWithCustomUI.ts`

```typescript
export async function newBookWithCustomUIAsync(context: CommandContext)
```

**PowerShell**: `bin/New-Excel.ps1`

処理内容：
- テンプレートアドイン（.xlam）から新規ブック作成
- CustomUI テンプレートのコピー
- .xlsm 形式で保存

### テンプレートファイル

拡張機能に付属するテンプレート：
```
excel/addin/excel-vba-addin.xlam
  ├── ModuleOpenVSCode.bas     # VS Code 連携用サンプルコード
  ├── Sheet1.cls
  ├── ThisWorkbook.cls
  └── customUI.xml, customUI14.xml
```

## エラーハンドリング

| エラー条件 | メッセージ | 対応 |
|---------|---------|------|
| Excel が起動していない | − | 自動的に新規 Excel プロセスを起動 |
| ファイル保存失敗 | ファイル保存に失敗しました | ディスク容量などを確認 |
| テンプレート読み込み失敗 | テンプレートが見つかりません | 拡張機能のインストール状態を確認 |

## セキュリティ

- ファイルの内容は検証されません
- ユーザーの入力値がファイル名として直接使用されます
- 既存ファイルが上書きされる可能性があります

## パフォーマンス

- Excel の起動: 1-3 秒（初回起動時）
- ファイル作成: < 1 秒

## 制限事項

- Excel 2010 以降が必要
- リボン定義は基本テンプレートのみ
- VBA コードは含まれません（CustomUI テンプレート版のみ）
