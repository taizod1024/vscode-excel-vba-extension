# Export Sheet as PNG 機能仕様書

## 概要

Excel ブックの指定シートの印刷範囲を PNG 画像にエクスポートする機能です。

## 機能概要

- シート名が `.png` で終わるシートを処理対象とします
- 各シートの印刷範囲を 1 枚の PNG 画像にエクスポートします
- 複数のシートを一括処理できます

## ユースケース

- セルの枠線やページ番号を表示せずに、Excelの特定範囲のデータを画像として保存したい
- ダッシュボードやレポートなどの資料として、Excelシートを PNG 画像に変換したい
- バージョン管理システムで画像ファイルとして Excel 情報を管理したい

## 入力仕様

### ファイル

- **対象**: Excel ファイル（.xlsx, .xlsm, .xlam）
- **前提条件**: 処理対象のシートが Excel で開かれている必要があります

### シード条件

1. **シート名**: `.png` で終わる名前
   - 例: `Dashboard.png`, `Report.png`, `Chart.png`

2. **印刷範囲**: 各シートに印刷範囲が設定されている必要があります
   - Excel でシートを選択 →「ページレイアウト」タブ → 「印刷範囲の設定」
   - 例: `$A$1:$E$15`

3. **表示設定（推奨）**: 画像に不要な要素を表示しない場合
   - グリッド線を非表示（「表示」タブ → グリッド線のチェックを外す）
   - ページ区切りプレビューで表示されるページ番号を非表示

## 出力仕様

### 出力フォルダ

```
{ブック名}_png/
```

- ブック名の末尾に `_png` を追加したフォルダが作成されます
- 例: `test.xlsx` → `test_png` フォルダ

### 出力ファイル

```
{シート名}
```

- ファイル形式: PNG（`.png` 拡張子）
- ファイル名: シート名そのもの（`.png` を含む）
- 例: `Dashboard.png`シート → `Dashboard.png` ファイル

### 出力内容

- 印刷範囲に指定された領域のみが PNG 画像として保存されます
- セルの値、フォーマット（背景色、文字色など）、図形が含まれます

## 処理仕様

### 処理フロー

1. ブックファイルを検出
2. Excel で開かれているブックを確認
3. ブック内の全シートをスキャン
4. シート名が `.png` で終わるシートを抽出
5. 各シートの印刷範囲を確認
6. 印刷範囲をクリップボードにコピー
7. クリップボードの内容を PNG 画像として保存
8. `_png` フォルダに保存

### エラーハンドリング

| エラー条件 | メッセージ | 対応 |
|---------|---------|------|
| Excel が起動していない | "Excel not running." | Excel を起動して、処理対象のファイルを開きます |
| ブックが開かれていない | "No workbook open." | Excel でブックを開きます |
| 印刷範囲が未設定 | "Sheet '{name}' has no print area defined." | Excel で印刷範囲を設定します |
| クリップボード取得失敗 | "Failed to get image from clipboard" | 処理をやり直します |

## 技術仕様

### マイン処理言語

- **メイン**: TypeScript (VS Code Extension)
- **サブ処理**: PowerShell (bin/Export-SheetAsPng.ps1)

### 依存関係

#### PowerShell

- `System.Drawing`: GDI+ 画像処理
- `System.Windows.Forms`: クリップボード操作
- Excel COM オブジェクト

#### TypeScript

- vscode: VS Code API
- child_process: PowerShell 実行

### 実行画面

```
Export Sheet as PNG
- File: test.xlsx
- Output: test_png
- Output: Export-SheetAsPng.ps1:
  - bookPath: C:\path\to\test.xlsx
  - imageOutputPath: C:\path\to\test_png
  - checking if workbook file exists
  - checking if workbook/add-in is open in Excel
  - Total sheets: 2
  - Sheets ending with .png: 1
  - Exporting: aaa.png
    - Print area: $A$1:$E$15
    - Copied to clipboard
    - Saved to: C:\path\to\test_png\aaa.png
  - Export complete
```

## セキュリティ考慮事項

- クリップボード経由で画像データを処理するため、セキュリティソフトが干渉する可能性があります
- 大規模な印刷範囲の場合、メモリ使用量が増加します

## パフォーマンス

- 印刷範囲のサイズに応じて処理時間が変動します
- 複数シートを処理する場合、シートごとに順次処理されます

## 制限事項

1. 複数ページの印刷範囲は 1 枚の PNG に統合されます
2. 印刷設定（拡大・縮小など）は反映されません
3. 動的コンテンツ（グラフなど）は画像化時点の状態で固定されます
4. シート保護されたシートも処理可能です

## 使用例

### 適用例

```
test.xlsx
├── Dashboard.png シート（印刷範囲: A1:G30）
├── Report.png シート（印刷範囲: A1:E100）
└── Data シート（処理対象外）

実行後:

test_xlsx_png/
├── Dashboard.png（画像ファイル）
└── Report.png（画像ファイル）
```

## 今後の拡張案

- [ ] 複数ページの出力時に各ページを別ファイルで保存
- [ ] 出力形式（JPEG, BMP など）のオプション指定
- [ ] 画像解像度の指定オプション
- [ ] 背景色の指定オプション
