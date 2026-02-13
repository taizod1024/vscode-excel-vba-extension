# Excel VBA Extension - 全機能仕様書

## 概要

Excel VBA Extension は VS Code から Excel VBA、CSV、CustomUI を開発・管理するための拡張機能です。

## 機能一覧

| # | 機能 | 説明 |
|----|------|------|
| 1 | New Excel Book | 新しい Excel ファイル（.xlsx）を作成 |
| 2 | New Excel Book with CustomUI | CustomUI テンプレート付きの Excel ファイル（.xlsm）を作成 |
| 3 | Open Excel Book | Excel ファイルを Excel で開く |
| 4 | Load VBA from Excel Book | Excel ブックから VBA ファイルを抽出 |
| 5 | Save VBA to Excel Book | VS Code の VBA ファイルを Excel ブックに保存 |
| 6 | Compare VBA with Excel Book | VS Code と Excel ブックのコード差分を表示 |
| 7 | Run VBA Sub at Cursor | カーソル位置の Sub プロシージャを Excel で実行 |
| 8 | Load CSV from Excel Book | Excel シートを CSV ファイルとしてエクスポート |
| 9 | Save CSV to Excel Book | CSV ファイルを Excel シートに保存 |
| 10 | Create URL Shortcut | クラウドホストの Excel ファイル用 URL ショートカットを作成 |
| 11 | Load CustomUI from Excel Book | Excel ブックから CustomUI XML をエクスポート |
| 12 | Save CustomUI to Excel Book | VS Code の CustomUI XML を Excel ブックに保存 |
| 13 | Export Sheet as PNG | シートを PNG 画像としてエクスポート |

## アーキテクチャ

```
VS Code Extension (TypeScript)
        ↓
   execPowerShell()
        ↓
   PowerShell Scripts
        ↓
Excel COM API
```

### メイン処理フロー

1. VS Code でコマンド実行
2. TypeScript が入力値を解析
3. 対応する PowerShell スクリプトを実行
4. PowerShell 内でファイル操作や Excel 処理
5. 結果を VS Code の出力チャネルに表示

## ファイル構成

```
src/
  commands/
    createUrlShortcut.ts
    loadCsv.ts
    loadCustomUI.ts
    loadVba.ts
    newBook.ts
    newBookWithCustomUI.ts
    openBook.ts
    runSub.ts
    saveCsv.ts
    saveCustomUI.ts
    saveVba.ts
    compareVba.ts
    exportSheetAsImage.ts
  utils/
    editorOperations.ts
    excelPath.ts
    execPowerShell.ts
    fileOperations.ts
    pathResolution.ts
    types.ts
    urlFile.ts
    vbValidation.ts

bin/
  Common.ps1                    # 共通関数
  Create-UrlShortcuts.ps1
  Load-CSV.ps1
  Load-CustomUI.ps1
  Load-VBA.ps1
  New-Excel.ps1
  Run-Sub.ps1
  Save-CSV.ps1
  Save-CustomUI.ps1
  Save-VBA.ps1
  Export-SheetAsImage.ps1
```

## 共通仕様

### エラーハンドリング

全コマンドで以下のエラーハンドリングが実装されています：

1. **例外キャッチ**
   - try-catch で全エラーをキャッチ
   - エラー内容を出力チャネルに記録
   - ユーザーに通知

2. **リソース管理**
   - COM オブジェクトは ReleaseComObject で解放
   - テンポラリファイルは自動削除

### ファイルパス解析

拡張機能は以下のファイルタイプを自動認識：

- `.xlsx`, `.xlsm`, `.xlam` - Excel ブック
- `.url` - URL ショートカット（クラウドホスト）
- `.csv` - CSV ファイル
- `.bas`, `.cls`, `.frm` - VBA モジュール
- `.xml` - CustomUI XML

選択ファイルのタイプに応じて、自動的に対応する Excel ブックを検出します。

### 出力フォルダ命名規則

```
{ブック名}_{種類}
```

| 種類 | フォルダ名 | 内容 |
|------|----------|------|
| VBA | `{name}_bas` | .bas, .cls, .frm ファイル |
| CSV | `{name}_csv` | .csv ファイル |
| CustomUI | `{name}_xml` | customUI.xml, customUI14.xml |
| PNG | `{name}_png` | .png 画像ファイル |

## ドキュメント構成

- [ALL_FEATURES_OVERVIEW.md](./ALL_FEATURES_OVERVIEW.md) - このドキュメント
- [NEW_BOOK_SPEC.md](./NEW_BOOK_SPEC.md) - New Excel Book の仕様書
- [LOAD_SAVE_VBA_SPEC.md](./LOAD_SAVE_VBA_SPEC.md) - VBA Load/Save/Compare の仕様書
- [LOAD_SAVE_CSV_SPEC.md](./LOAD_SAVE_CSV_SPEC.md) - CSV Load/Save の仕様書
- [LOAD_SAVE_CUSTOMUI_SPEC.md](./LOAD_SAVE_CUSTOMUI_SPEC.md) - CustomUI Load/Save の仕様書
- [EXPORT_SHEET_AS_PNG_SPEC.md](./EXPORT_SHEET_AS_PNG_SPEC.md) - Export Sheet as PNG の仕様書
- [CLOUD_SUPPORT_SPEC.md](./CLOUD_SUPPORT_SPEC.md) - URL Shortcut とクラウドサポートの仕様書
- [COMMAND_REFERENCE.md](./COMMAND_REFERENCE.md) - コマンドリファレンス
