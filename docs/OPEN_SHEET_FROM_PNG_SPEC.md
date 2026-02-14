# Open Sheet from PNG 機能仕様書

## 概要

PNG 画像ファイルから元の Excel ワークブックを開き、対応するシートをアクティベートする機能です。

## 機能概要

- PNG ファイル名をシート名として使用
- フォルダ構造から対応する Excel ファイルを自動検派出
- Excel を開いて指定されたシートをアクティベート

## ユースケース

- PNG 画像をエクスポートした元の Excel シートを再度編集したい
- PNG 画像をクリックするだけで元データが開くようにしたい
- 画像レビュー中に元のデータを確認したい

## 入力仕様

### ファイル

- **対象**: PNG ファイル（`.png`）
- **前提条件**: Export Sheet as PNG で生成された PNG ファイル

### フォルダ構造

```
{作業ディレクトリ}/
├── {ブック名}.xlsm          ← 開く Excel ファイル
└── {ブック名}_xlsm/         ← PNG を含むフォルダ
    └── png/
        └── {シート名}.png   ← クリックするファイル
```

### ファイル名ルール

- **シート名**: PNG ファイル名そのもの（拡張子を含む）
  - 例: `Dashboard.png` ファイル → `Dashboard.png` という名前のシートを検索
  - 例: `Report.png` ファイル → `Report.png` という名前のシートを検索

### 検出ロジック

1. PNG ファイル名からシート名を抽出（拡張子を含める）
2. パスから親フォルダを走査
   - `{name}_xlsx/png/Sheet.png` → `{name}.xlsx` を検索
   - `{name}_xlsm/png/Sheet.png` → `{name}.xlsm` を検索
   - `{name}_xlam/png/Sheet.png` → `{name}.xlam` を検索

## 出力仕様

### 処理結果

- Excel アプリケーションで対応するシートがアクティベートされた状態で表示

### メッセージ

| 状態         | 出力                                                           |
| ------------ | -------------------------------------------------------------- |
| 成功         | `Sheet '{sheet}' selected successfully`                        |
| ファイルなし | `Excel file not found: {path}`                                 |
| シートなし   | `Sheet '{sheet}' not found in workbook. Available sheets: ...` |
| 無効なパス   | `Invalid folder structure. Expected format: filename_ext`      |

## 処理仕様

### 処理フロー

1. PNG ファイルパスを取得
2. PNG ファイルが有効か確認（`.png` 拡張子）
3. PNG ファイル名をシート名として抽出
4. フォルダ構造からから Excel ファイルパスを構築
5. Excel ファイルが存在するか確認
6. PowerShell スクリプトを実行
   - Excel を起動（または既存インスタンスを取得）
   - Excel ファイルを開く
   - シート名で検索してアクティベート
   - Excel ウィンドウをアクティブ化

### エラーハンドリング

| エラー条件         | メッセージ                                                      | 対応                                          |
| ------------------ | --------------------------------------------------------------- | --------------------------------------------- |
| PNG でない         | "`File is not a PNG file.`"                                     | PNG ファイルを選択してください                |
| パス無効           | "`Invalid folder structure. Expected format: filename_ext`"     | PNG ファイルがフォルダ構造にないか確認        |
| Excel ファイルなし | "`Excel file not found: {path}`"                                | Excel ファイルを作成してください              |
| シートなし         | "`Sheet '{name}' not found in workbook. Available sheets: ...`" | シート名が PNG ファイル名と一致しているか確認 |

## 技術仕様

### 処理言語

- **メイン**: TypeScript (VS Code Extension)
  - ファイルパス解析
  - PowerShell スクリプト実行
- **サブ処理**: PowerShell (bin/Open-SheetFromPng.ps1)
  - Excel COM オートメーション
  - シート検索・アクティベート

### 依存関係

#### PowerShell

- Excel COM オブジェクト
- `WScript.Shell`: ウィンドウアクティベーション

#### TypeScript

- Node.js `path` モジュール: パス解析
- Node.js `fs` モジュール: ファイル存在確認
- vscode: 出力チャネル表示
- child_process: PowerShell 実行

### 実行画面

```
## Open Sheet from PNG
- file: aaa.png
- path: c:\Users\taizo\Downloads\aaa7_xlsm\png\aaa.png
- PNG directory: c:\Users\taizo\Downloads\aaa7_xlsm\png
- XLSX folder name: aaa7_xlsm
- Excel file path: c:\Users\taizo\Downloads\aaa7.xlsm
- Sheet name: aaa.png
Sheet 'aaa.png' selected successfully
```

## ワークフロー（推奨）

### Export → Review → Edit

```
1. Excel で「aaa.png」という名前のシートを作成
   ↓
2. Export Sheet as PNG を実行
   → aaa7_xlsm/png/aaa.png が生成される
   ↓
3. PNG ファイルをレビュー
   ↓
4. Open Sheet from PNG を実行
   → Excel の「aaa.png」シートが開く
   ↓
5. Excel でシートを編集
   ↓
6. 改めて Export Sheet as PNG を実行
```

## 制限事項

1. PNG ファイル名がシート名と完全一致する必要があります
2. PNG ファイルが フォルダ構造 `{name}_拡張子/png/` 内にある必要があります
3. Excel ファイルが同じ親ディレクトリにある必要があります
4. シート名に特殊文字が含まれる場合は、PNG ファイル名も同じ特殊文字を使用する必要があります

## セキュリティ考慮事項

- Excel ファイルは自動的に開かれるため、マクロ警告が表示される場合があります
- 信頼できるファイルのみを実行してください

## パフォーマンス

- Excel アプリケーションの起動有無により1-3秒程度の時間がかかります
- Excel が既に起動している場合は更に高速です

## 関連機能

- [Export Sheet as PNG](./EXPORT_SHEET_AS_PNG_SPEC.md) - PNG へのエクスポート
- [Open Excel Book](./COMMAND_REFERENCE.md) - Excel ファイルを直接開く
