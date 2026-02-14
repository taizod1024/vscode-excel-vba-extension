# コマンドリファレンス

## クイックスタート

### VBA 開発の基本フロー

```
1. 「New Excel Book with CustomUI」で新しいファイルを作成
   ↓
2. 「Load VBA from Excel Book」で VBA を VS Code に抽出
   ↓
3. VS Code で VBA ファイルを編集
   ↓
4. 「Save VBA to Excel Book」で Excel に保存
   ↓
5. 「Run VBA Sub at Cursor」で実行 & デバッグ
```

### CSV シート管理の基本フロー

```
1. 「Load CSV from Excel Book」でシートを CSV に変換
   ↓
2. VS Code で CSV を編集（バージョン管理可）
   ↓
3. 「Save CSV to Excel Book」で Excel シートに保存
```

### CustomUI 開発の基本フロー

```
1. 「Load CustomUI from Excel Book」で XML をエクスポート
   ↓
2. VS Code で XML を編集（IntelliSense 活用）
   ↓
3. 「Save CustomUI to Excel Book」で保存
   ↓ ※ Excel を閉じておくこと
```

---

## コマンド詳細

### 1. New Excel Book

**コマンド ID**: `excel-vba.newBook`

**実行方法**

- File → New File → New Excel Book
- コマンドパレット → "New Excel Book"

**入力**

- ファイル名（拡張子なし）

**出力**

- `{name}.xlsx`（ワークスペース直下）

**処理時間**: < 5 秒

**参照**: [NEW_BOOK_SPEC.md](./NEW_BOOK_SPEC.md#1-new-excel-book)

---

### 2. New Excel Book with CustomUI as Macro

**コマンド ID**: `excel-vba.newBookWithCustomUI`

**実行方法**

- コマンドパレット → "New Excel Book with CustomUI as Macro"

**入力**

- ファイル名（拡張子なし）

**出力**

- `{name}.xlsm`（ワークスペース直下）
- CustomUI テンプレート付き

**処理時間**: < 5 秒

**参照**: [NEW_BOOK_SPEC.md](./NEW_BOOK_SPEC.md#2-new-excel-book-with-customui)

---

### 3. Open Excel Book

**コマンド ID**: `excel-vba.openBook`

**実行方法**

- ファイルを右クリック → Open Excel Book
- エディタタイトルアイコン

**対象ファイル**

- `.xlsx`, `.xlsm`, `.xlam`, `.url`, `.csv`, `.bas`, `.cls`, `.frm`, `.xml`

**動作**

- Excel で対象ファイルを開く
- VBA/CSV/XML ファイルの場合、対応する Excel ファイルを検出・開く

**処理時間**: 1-5 秒

**参照**: [CLOUD_SUPPORT_SPEC.md](./CLOUD_SUPPORT_SPEC.md#2-open-excel-book)

---

### 4. Load VBA from Excel Book

**コマンド ID**: `excel-vba.loadVba`

**実行方法**

- ブックファイルを右クリック → Load VBA from Excel Book
- エディタタイトルアイコン

**前提条件**

- Excel でブックが開かれている

**出力**

- `{name}.bas/` フォルダ
- `.bas`, `.cls`, `.frm` ファイル

**処理時間**: 1-5 秒（コード量による）

**参照**: [LOAD_SAVE_VBA_SPEC.md](./LOAD_SAVE_VBA_SPEC.md#1-load-vba-from-excel-book)

---

### 5. Save VBA to Excel Book

**コマンド ID**: `excel-vba.saveVba`

**実行方法**

- ブックファイルを右クリック → Save VBA to Excel Book
- エディタタイトルアイコン

**前提条件**

- Excel でブックが開かれている
- `{name}.bas/` フォルダが存在

**動作**

- `.bas` フォルダ内のファイルを読み込み
- 既存 VBA コンポーネントをすべて削除・再作成
- Excel ブックに保存

**処理時間**: 1-5 秒（コード量による）

**参照**: [LOAD_SAVE_VBA_SPEC.md](./LOAD_SAVE_VBA_SPEC.md#2-save-vba-to-excel-book)

---

### 6. Compare VBA with Excel Book

**コマンド ID**: `excel-vba.compareVba`

**実行方法**

- ブックファイルを右クリック → Compare VBA with Excel Book
- エディタタイトルアイコン

**前提条件**

- Excel でブックが開かれている
- `{name}.bas/` フォルダが存在

**動作**

- Excel のコードを一時ファイルに書き込み
- VS Code で Diff ビュー表示
- 左窒: Excel, 右窒: VS Code

**処理時間**: < 1 秒

**参照**: [LOAD_SAVE_VBA_SPEC.md](./LOAD_SAVE_VBA_SPEC.md#3-compare-vba-with-excel-book)

---

### 7. Run VBA Sub at Cursor

**コマンド ID**: `excel-vba.runSub`

**実行方法**

- VBA ファイル内で実行 → Run VBA Sub at Cursor
- エディタタイトルアイコン

**前提条件**

- VBA ファイルを開いている
- Excel でブックが開かれている
- カーソルが Sub 内

**動作**

1. VBA ファイルを Save
2. カーソル位置の Sub を特定
3. Excel で Sub を実行
4. 結果を表示

**処理時間**: < 1 秒 + 実行時間

**参照**: [LOAD_SAVE_VBA_SPEC.md](./LOAD_SAVE_VBA_SPEC.md#4-run-vba-sub-at-cursor)

---

### 8. Load CSV from Excel Book

**コマンド ID**: `excel-vba.loadCsv`

**実行方法**

- ブックファイルを右クリック → Load CSV from Excel Book
- エディタタイトルアイコン

**前提条件**

- Excel でブックが開かれている
- シート名が「.csv」で終わるシートが存在

**出力**

- `{name}.csv/` フォルダ
- 各シート = CSV ファイル

**処理時間**: 1-3 秒（シート数による）

**参照**: [LOAD_SAVE_CSV_SPEC.md](./LOAD_SAVE_CSV_SPEC.md#1-load-csv-from-excel-book)

---

### 9. Save CSV to Excel Book

**コマンド ID**: `excel-vba.saveCsv`

**実行方法**

- ブックファイルを右クリック → Save CSV to Excel Book
- エディタタイトルアイコン

**前提条件**

- Excel でブックが開かれている
- `{name}.csv/` フォルダが存在

**動作**

- `.csv` フォルダ内の CSV を読み込み
- 既存シートを削除
- 新しいシートを作成・データを書き込み
- ブック保存

**処理時間**: 1-3 秒（データ量による）

**参照**: [LOAD_SAVE_CSV_SPEC.md](./LOAD_SAVE_CSV_SPEC.md#2-save-csv-to-excel-book)

---

### 10. Create URL Shortcut

**コマンド ID**: `excel-vba.createUrlShortcut`

**実行方法**

- コマンドパレット → "Create URL Shortcut"

**前提条件**

- OneDrive / SharePoint のファイルが Excel で開かれている（複数可）

**出力**

- `{name}.url` ファイル（ワークスペース直下）

**動作**

1. Excel で開かれている全ブックをスキャン
2. クラウドホストの URL を抽出
3. .url ショートカット作成

**処理時間**: < 1 秒

**参照**: [CLOUD_SUPPORT_SPEC.md](./CLOUD_SUPPORT_SPEC.md#1-create-url-shortcut)

---

### 11. Load CustomUI from Excel Book

**コマンド ID**: `excel-vba.loadCustomUI`

**実行方法**

- アドインファイル（.xlam）を右クリック → Load CustomUI from Excel Book
- エディタタイトルアイコン

**前提条件**

- Excel でアドインが開かれている

**出力**

- `{name}.xml/` フォルダ
- `customUI.xml`, `customUI14.xml`

**処理時間**: < 1 秒

**参照**: [LOAD_SAVE_CUSTOMUI_SPEC.md](./LOAD_SAVE_CUSTOMUI_SPEC.md#1-load-customui-from-excel-book)

---

### 12. Save CustomUI to Excel Book

**コマンド ID**: `excel-vba.saveCustomUI`

**実行方法**

- アドインファイルを右クリック → Save CustomUI to Excel Book
- エディタタイトルアイコン

**前提条件**

- Excel を閉じている（重要）
- `{name}.xml/` フォルダが存在
- XML ファイルが編集済み

**動作**

1. アドインをZIP として開く
2. customUI パーツを更新
3. アドイン再構築
4. 保存

**処理時間**: 1-3 秒

**参照**: [LOAD_SAVE_CUSTOMUI_SPEC.md](./LOAD_SAVE_CUSTOMUI_SPEC.md#2-save-customui-to-excel-book)

---

### 13. Export Sheets as PNG

**コマンド ID**: `excel-vba.exportSheetsAsPng`

**実行方法**

- ブックファイルを右クリック → Export Sheet as PNG
- エディタタイトルアイコン

**前提条件**

- Excel でブックが開かれている
- シート名が「.png」で終わるシートが存在
- 各シートに印刷範囲が設定されている

**出力**

- `{name}_png/` フォルダ
- 各シート = PNG ファイル

**処理時間**: 1-3 秒（シート数による）

**参照**: [EXPORT_SHEET_AS_PNG_SPEC.md](./EXPORT_SHEET_AS_PNG_SPEC.md)

---

### 14. Open Sheet from PNG

**コマンド ID**: `excel-vba.openSheetFromPng`

**実行方法**

- PNG ファイルを右クリック → Open Sheet from PNG
- エディタタイトルアイコン

**対象ファイル**

- `.png`（Export Sheet as PNG で作成した PNG ファイル）

**前提条件**

- PNG ファイルが `{name}_拡張子/png/` フォルダ内に存在
- PNG ファイル名をシート名として使用（例: `aaa.png` → シート名 `aaa.png`）
- 対応する Excel ファイルが同じディレクトリまたは親ディレクトリに存在

**動作**

- PNG ファイル名からシート名を抽出
- フォルダ構造から対応する Excel ファイルを特定
- Excel でファイルを開く
- 指定されたシート名のシートをアクティベート

**処理時間**: 1-3 秒

**参照**: [OPEN_SHEET_FROM_PNG_SPEC.md](./OPEN_SHEET_FROM_PNG_SPEC.md)

---

## コマンド早見表

| #   | コマンド ID         | 説明                       | キー |
| --- | ------------------- | -------------------------- | ---- |
| 1   | newBook             | 新しい Excel ファイル作成  | −    |
| 2   | newBookWithCustomUI | CustomUI as Macro 付き作成 | −    |
| 3   | openBook            | Excel で開く               | −    |
| 4   | loadVba             | VBA 抽出                   | −    |
| 5   | saveVba             | VBA 保存                   | −    |
| 6   | compareVba          | VBA 比較                   | −    |
| 7   | runSub              | Sub 実行                   | −    |
| 8   | loadCsv             | CSV 抽出                   | −    |
| 9   | saveCsv             | CSV 保存                   | −    |
| 10  | createUrlShortcut   | URL ショートカット作成     | −    |
| 11  | loadCustomUI        | CustomUI 抽出              | −    |
| 12  | saveCustomUI        | CustomUI 保存              | −    |
| 13  | exportSheetsAsPng   | PNG エクスポート（複数）   | −    |
| 14  | openSheetFromPng    | PNG から Sheet を開く      | −    |

---

## キーボードショートカット

現在、カスタムショートカットは定義されていません。
VS Code の設定から任意に追加できます。

例:

```json
{
  "key": "ctrl+alt+l",
  "command": "excel-vba.loadVba"
}
```
