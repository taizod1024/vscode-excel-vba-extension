# Export Sheet as PNG 実装ガイド

## ファイル構成

```
src/commands/exportSheetAsImage.ts    # メイン処理（TypeScript）
bin/Export-SheetAsImage.ps1           # PowerShell スクリプト
```

## 実装詳細

### 1. TypeScript 部分 (src/commands/exportSheetAsImage.ts)

#### 機能

VS Code のコマンド実行時に呼び出される非同期関数です。

#### 実装フロー

```
exportSheetAsPngAsync
  ↓
パス解析 (マクロパス → ファイル名、ディレクトリ)
  ↓
出力パス生成 ({ブック名}_png)
  ↓
PowerShell スクリプト実行
  ↓
結果出力 (出力チャネル)
```

#### コード例

```typescript
export async function exportSheetAsPngAsync(bookPath: string, context: CommandContext) {
  // ファイルパス抽出
  const bookFileName = path.parse(bookPath).name;
  const bookDir = path.dirname(bookPath);
  const pngDir = path.join(bookDir, `${bookFileName}.png`);

  // PowerShell スクリプト実行
  const result = execPowerShell(scriptPath, [bookPath, pngDir]);
}
```

### 2. PowerShell 部分 (bin/Export-SheetAsImage.ps1)

#### 処理ステップ

1. **初期化**
   - スクリプト名出力
   - 入力パラメータ検証

2. **ブック確認**
   ```powershell
   $macroInfo = Get-BookInfo $bookPath
   $result = Find-VBProject $excel $macroInfo.ResolvedPath $macroInfo.IsAddIn
   $workbook = $result.Workbook
   ```

3. **出力フォルダ作成**
   ```powershell
   if (Test-Path $pngOutputPath) {
       Remove-Item $pngOutputPath -Recurse -Force
   }
   New-Item -ItemType Directory -Force -Path $pngOutputPath | Out-Null
   ```

4. **シートスキャン**
   - 全シート数カウント
   - `.png` で終わるシートを識別

5. **PNG エクスポート処理**

   各シートについて以下を実行：
   - シート選択
   - 印刷範囲確認
   - 範囲をクリップボードにコピー
   - GDI+ で画像取得
   - PNG ファイル保存

#### キー処理

**クリップボード → PNG 変換**

```powershell
# GDI+ アセンブリ読み込み
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

# クリップボード取得
$image = [System.Windows.Forms.Clipboard]::GetImage()

# PNG 保存
$image.Save($outputFile, [System.Drawing.Imaging.ImageFormat]::Png)
```

## デバッグ情報

### PowerShell スクリプト内の出力

```
- bookPath: {ファイルパス}
- pngOutputPath: {出力フォルダパス}
- checking if workbook file exists
- checking if workbook/add-in is open in Excel
- Total sheets: {シート数}
- Sheets ending with .png: {処理対象シート数}
  - Exporting: {シート名}
    - Print area: {印刷範囲}
    - Copied to clipboard
    - Saved to: {出力ファイルパス}
- Export complete
```

### ログレベル

- **情報**: デバッグ出力、処理進捗
- **エラー**: 処理失敗時のエラーメッセージ

## トラブルシューティング

### 症状 1: "Excel not running"

**原因**: Excel が起動していない

**対応**:
- Excel を起動
- 処理対象のファイルを開く
- コマンドを再実行

### 症状 2: "No book open."

**原因**: ブックが Excel で開かれていない

**対応**:
- Excel でブックを開く
- コマンドを再実行

### 症状 3: "has no print area defined"

**原因**: シートに印刷範囲が設定されていない

**対応**:
- Excel でシートを選択
- 「ページレイアウト」タブで「印刷範囲の設定」を実行
- コマンドを再実行

### 症状 4: 出力画像に枠線やページ番号が表示される

**原因**: シートの表示設定が有効になっている

**対応**:
- Excel でシートを選択
- 「表示」タブでグリッド線のチェックを外す
- コマンドを再実行

## パフォーマンスOptimization

### 推奨事項

1. **打ち続きではなく、明確に設定された印刷範囲を使用**
   - 小さな範囲ほど処理が高速

2. **シート数の削減**
   - `.png` で終わるシートのみを処理対象にする

3. **Excel メモリ管理**
   - 大規模な画像は事前に縮小を検討

## テスト方法

### テストケース 1: 基本動作

1. Excel で `test.xlsx` を開く
2. `Dashboard.png` シート作成（印刷範囲: A1:E15）
3. VS Code で「Export Sheet as PNG」を実行
4. `test_png` フォルダに `Dashboard.png` が作成されることを確認

### テストケース 2: 複数シート

1. Excel で `test.xlsx` を開く
2. `Report.png` シート作成（印刷範囲: A1:K50）
3. VS Code で「Export Sheet as PNG」を実行
4. `test_png` フォルダに 2 つの PNG ファイルが作成されることを確認

### テストケース 3: エラーハンドリング

1. 印刷範囲未設定のシートでテスト
2. Excel 未起動状態でテスト
3. エラーメッセージが正しく表示されることを確認

## パッケージング

### ファイル構成の確認

```
build/main.js                          # ビルド済みメイン（自動生成）
bin/Export-SheetAsImage.ps1            # PowerShell スクリプト
package.json                           # コマンド定義
```

### ビルド

```bash
npm run esbuild
```
