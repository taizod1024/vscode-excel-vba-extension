# CSV 機能仕様書 (Load / Save)

## 概要

Excel シートを CSV ファイルとして抽出・保存する機能です。

## 機能一覧

1. **Load CSV from Excel Book** - Excel シートを CSV ファイルとしてエクスポート
2. **Save CSV to Excel Book** - CSV ファイルを Excel シートに保存

## 1. Load CSV from Excel Book

### 概要

Excel ブック内で名前が「.csv」で終わるシートを CSV ファイルとしてエクスポートします。

### 入力仕様

**対象ファイル**
- Excel ブック: .xlsm, .xlam, .xlsx
- URL ショートカット: .url（クラウドホスト）

**対象シート**
- シート名が「.csv」で終わるシート
- 例: `data.csv`, `output.csv`, `2024.csv`

### 出力仕様

**出力フォルダ**
```
{ブック名}_csv/
  ├── data.csv
  ├── output.csv
  ├── 2024.csv
  └── ...
```

**ファイル形式**
- 形式: CSV（RFC 4180）
- 文字エンコーディング: UTF-8（BOM なし）
- 改行コード: LF（Unix 形式）
- 区切り文字: カンマ（,）
- クォート文字: ダブルクォート（"）

### 処理フロー

1. Excel でブックを開く（必須）
2. VS Code でブックファイルを選択
3. 「Load CSV from Excel Book」を実行
4. CSV エクスポート処理
   - 該当シートを検索
   - 各シートのセルデータを読み込み
   - CSV 形式に変換
   - ファイル保存
5. フォルダが作成され、CSV ファイルが表示

### シート検索ロジック

```
for each sheet in workbook.sheets:
  if sheet.name ends with ".csv":
    export sheet
```

### CSV 変換ロジック

**セル値の処理**
- テキスト値: そのまま出力
- 数値: 文字列に変換
- 論理値: TRUE/FALSE
- 日付: デフォルト形式で出力
- 空白セル: 空の値

**カンマ含む値**
```
値に「,」を含む場合:
  "value, with, comma"
```

**ダブルクォート含む値**
```
値に「"」を含む場合:
  "value ""with"" quote"
```

### 実装詳細

**メイン処理**: `src/commands/loadCsv.ts`

```typescript
export async function loadCsvAsync(macroPath: string, context: CommandContext)
```

**PowerShell**: `bin/Load-CSV.ps1`

処理内容：
1. VB プロジェクト検索
2. 全シート列挙
3. シート名が「.csv」で終わるものを抽出
4. 各シートを CSV に変換
5. ファイル保存

### 自動表示

CSV ファイル作成後、最初の CSV ファイルが自動で開きます。

## 2. Save CSV to Excel Book

### 概要

VS Code で編集した CSV ファイルを Excel シートに保存します。

### 入力仕様

**入力ファイル**
```
{ブック名}_csv/
  ├── data.csv
  ├── output.csv
  └── ...
```

### 処理フロー

1. Excel でブックを開く（必須）
2. VS Code でブックファイル（または CSV ファイル）を選択
3. 「Save CSV to Excel Book」を実行
4. CSV インポート処理
   - _csv フォルダ内の全 CSV ファイルを読み込み
   - 既存シートを削除
   - 新しいシートを作成
   - CSV データをシートに書き込み
   - ブック保存

### シート作成ロジック

```
for each csv file in {name}_csv:
  sheet_name = csv_filename (without .csv)
  create sheet with sheet_name
  write csv data to sheet
```

### CSV パース仕様

**RFC 4180 準拠**

```
標準 CSV 形式で以下を処理：
- ヘッダ行の有無は自動判定
- カンマ区切り
- ダブルクォートでエスケープ
- CRLF または LF 対応
```

### エラー時の動作

| エラー条件 | 対応 |
|---------|------|
| CSV パースエラー | エラー通知、処理中止 |
| シート作成失敗 | エラー通知、処理中止 |
| セル書き込み失敗 | 部分的に処理、警告表示 |

### 実装詳細

**メイン処理**: `src/commands/saveCsv.ts`

```typescript
export async function saveCsvAsync(macroPath: string, context: CommandContext)
```

**PowerShell**: `bin/Save-CSV.ps1`

処理内容：
1. VB プロジェクト検索
2. CSV フォルダの存在確認
3. 既存シート削除
4. CSV ファイル読み込み
5. シート作成
6. データ書き込み
7. ブック保存

## セル型の処理

### Load CSV

すべてのセル値が**テキスト**として Excel に保存されます。

```
CSV: 123
Excel セル: "123"（テキスト）
```

### Save CSV

Excel のセル型を考慮して出力：

```
テキスト "123" → CSV: 123
数値 123 → CSV: 123
日付 2024-01-01 → CSV: 45293（シリアル値）
論理値 TRUE → CSV: TRUE
```

## 文字エンコーディング

**CSV 出力**: UTF-8（BOM なし）
**CSV 入力**: UTF-8, Shift_JIS, ASCII を自動判定

## パフォーマンス

| 操作 | シートサイズ | 処理時間 |
|------|----------|--------|
| Load CSV | 10K セル | < 1 秒 |
| Load CSV | 100K セル | 1-3 秒 |
| Save CSV | 10K セル | < 1 秒 |
| Save CSV | 100K セル | 1-3 秒 |

## 制限事項

1. セルの書式（日付形式など）は保持されません
2. 複数シート間の参照式は処理されません
3. 画像、グラフ、図形は処理されません
4. 条件付き書式は処理されません
5. セルのコメントは処理されません
6. シート保護設定は処理されません（パスワードなし）
