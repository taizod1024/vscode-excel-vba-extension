# VBA 機能仕様書 (Load / Save / Compare / Run Sub)

## 概要

Excel VBA コードと VS Code を連携させる機能です。

## 機能一覧

1. **Load VBA from Excel Book** - Excel ブックから VBA ファイルを抽出
2. **Save VBA to Excel Book** - VS Code の VBA ファイルを Excel ブックに保存
3. **Compare VBA with Excel Book** - コード差分を表示
4. **Run VBA Sub at Cursor** - Sub プロシージャを実行

## 1. Load VBA from Excel Book

### 概要

Excel ブックに含まれる VBA コード全体を VS Code で編集可能なファイルに抽出します。

### 入力仕様

**対象ファイル**
- Excel ブック: .xlsm, .xlam, .xlsx（マクロ有効）
- URL ショートカット: .url（クラウドホスト）

### 出力仕様

**出力フォルダ**
```
{ブック名}.bas/
  ├── StandardModule1.bas
  ├── StandardModule2.bas
  ├── ClassModule1.cls
  ├── Form1.frm
  ├── Sheet1.cls               # Sheet モジュール
  └── ThisWorkbook.cls
```

**ファイル形式**
- 拡張子: .bas（標準モジュール）、.cls（クラス）、.frm（ユーザーフォーム）
- 文字エンコーディング: Shift_JIS (日本語対応)

### 処理フロー

1. Excel でブックを開く（必須）
2. VS Code でブックファイルを選択
3. 「Load VBA from Excel Book」を実行
4. VBA コード抽出処理
   - 全 VB コンポーネントを列挙
   - 各コンポーネントのコードを抽出
   - メタデータ（Version, Attributes）を削除
   - ファイル保存
5. フォルダが作成され、ファイルが表示

### 実装詳細

**メイン処理**: `src/commands/loadVba.ts`

```typescript
export async function loadVbaAsync(macroPath: string, context: CommandContext)
```

**PowerShell**: `bin/Load-VBA.ps1`

処理内容：
- VB プロジェクト検索
- Document モジュール（Sheet, ThisWorkbook）の処理
- Standard/Class モジュールの抽出
- ユーザーフォーム（.frx ファイル）の処理

### ドキュメントモジュール処理

Document モジュール（Sheet, ThisWorkbook）は以下の処理を実施：

1. Attribute, VERSION, Begin, End 行を削除
2. 実装コードのみを抽出
3. Option Explicit から実装コード開始まで

### エラーハンドリング

| エラー条件 | メッセージ | 対応 |
|---------|---------|------|
| Excel が起動していない | "Excel not running." | Excel を起動 |
| ブックが開かれていない | "No workbook open." | Excel でブックを開く |
| 無効なVBA プロジェクト | VBA object model access denied | Trust Center でアクセス許可 |

## 2. Save VBA to Excel Book

### 概要

VS Code で編集した VBA ファイルを Excel ブックに保存します。

### 入力仕様

**入力ファイル**
```
{ブック名}_bas/
  ├── Module1.bas
  ├── Class1.cls
  ├── Form1.frm
  └── ...
```

### 処理フロー

1. Excel でブックを開く（必須）
2. VS Code でブックファイルを選択
3. 「Save VBA to Excel Book」を実行
4. 保存処理
   - 既存 VBA コンポーネントをすべて削除
   - .bas フォルダ内のファイルを読み込み
   - VB プロジェクトに新しいコンポーネントを追加
   - コード内容を設定
   - 保存完了

### VB_Name 属性検証

保存時に VB_Name 属性の検証が実施されます：

**規則**
- ファイル名（拡張子なし）= VB_Name の値
- コンパイルエラーを事前に検出

**例**
```
Module1.bas
  → Attribute VB_Name = "Module1"

MyClass.cls
  → Attribute VB_Name = "MyClass"
```

### 実装詳細

**メイン処理**: `src/commands/saveVba.ts`

```typescript
export async function saveVbaAsync(macroPath: string, context: CommandContext)
```

**PowerShell**: `bin/Save-VBA.ps1`

処理内容：
1. VB プロジェクト検索
2. 古いコンポーネント削除
3. 新しいコンポーネント追加
4. コード設定
5. ファイル保存

## 3. Compare VBA with Excel Book

### 概要

VS Code のコードと Excel ブックのコードを比較し、差分を表示します。

### 処理フロー

1. Excel でブックを開く（必須）
2. VS Code でブックファイルを選択
3. 「Compare VBA with Excel Book」を実行
4. 比較処理
   - Excel のコードを読み込み
   - VS Code のコードと比較
   - 差分ファイルを作成
5. Diff ビューで差分を表示

### 出力形式

Diff ビュー（左:Excel, 右:VS Code）

### 実装詳細

**メイン処理**: `src/commands/compareVba.ts`

```typescript
export async function compareVbaAsync(macroPath: string, context: CommandContext)
```

利用方法：
- 左側: Excel ブックのコード（編集不可）
- 右側: VS Code のコード（編集可能）
- 差分が色分け表示される

## 4. Run VBA Sub at Cursor

### 概要

VS Code のカーソル位置に在る Sub プロシージャを Excel で実行します。

### 処理フロー

1. VS Code で VBA ファイルを開く
2. 実行したい Sub プロシージャ内にカーソルを配置
3. 「Run VBA Sub at Cursor」を実行
4. 自動処理
   - VBA ファイルを Excel に保存
   - カーソル位置の Sub を特定
   - Excel でマクロを実行
5. 実行結果が返される

### Sub 特定ロジック

カーソル位置から以下のロジックで Sub を特定：

1. カーソル行から上方向に「Sub」キーワードを検索
2. 最も近い「Sub」を対象 Sub として特定
3. Sub の引数、属性（Private/Public）を読み込み
4. 実行

### 実装詳細

**メイン処理**: `src/commands/runSub.ts`

```typescript
export async function runSubAsync(macroPath: string, context: CommandContext)
```

**処理順序**
1. Save VBA（自動）→ Run Sub

### エラーハンドリング

| エラー条件 | メッセージ | 対応 |
|---------|---------|------|
| Sub が見つからない | Sub procedure not found | カーソルが Sub 内に在ることを確認 |
| VBA コンパイルエラー | Compile error | VBA コードの構文を確認 |
| 実行時エラー | Runtime error: ... | プログラムロジックを確認 |

## 文字エンコーディング

VBA ファイルは Shift_JIS 文字コードで保存されます。

**対応文字**
- ASCII
- 日本語（ひらがな、カタカナ、漢字）
- 記号

## パフォーマンス

| 操作 | 処理時間 |
|------|--------|
| Load VBA | 1-5 秒（コード量による） |
| Save VBA | 1-5 秒（コード量による） |
| Compare | < 1 秒 |
| Run Sub | < 1 秒（マクロ実行時間は別） |

## 制限事項

1. Document モジュール（Sheet, ThisWorkbook）は部分的に処理
2. ユーザーフォーム（.frm）と対応ファイル（.frx）は一体で処理
3. .frx ファイルを削除すると .frm は保存できません
4. マクロセキュリティが有効な場合、実行権限が必要
