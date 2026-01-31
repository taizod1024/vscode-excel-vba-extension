[ENGLISH](#excel-vba-extension-en) | [日本語](#excel-vba-extension-jp)

---

## Excel VBA Extension EN

A VS Code extension for developing Excel VBA.

**Key Benefits:**

- Edit VBA files in VS Code
- Run macros from VS Code
- Integration with version control systems
- Leverage AI assistants like GitHub Copilot

![excel-vba](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba.png?raw=true)

## Features

| Feature                            | Description                                                           |
| ---------------------------------- | --------------------------------------------------------------------- |
| **Open Excel Macro**               | Open Excel files from Explorer View.                                  |
| **Load VBA from Excel Macro**      | Extract VBA files (.bas, .cls, .frm) from macro files (.xlsm, .xlam). |
| **Save VBA to Excel Macro**        | Save VBA files edited in VS Code to macro files.                      |
| **Compare VBA with Excel Macro**   | Compare code edited in VS Code with the state of macro files.         |
| **Run VBA Sub at Cursor**          | Execute the Sub procedure at cursor position in Excel.                |
| **Load CustomUI from Excel Macro** | Export CustomUI (customUI.xml, customUI14.xml) from macro files.      |
| **Save CustomUI to Excel Macro**   | Save CustomUI edited in VS Code to macro files.                       |

## Setup

### Macro Security Settings

Configure settings to allow the extension to access Excel VBA.

1. Open Excel
2. Select **File** → **Options** → **Trust Center**
3. Click **Trust Center Settings**
4. In the **Macro Settings** section, verify:
   - Check **Trust access to the VBA project object model**
5. Click **OK** to save settings

## Usage

Right-click a macro file in Explorer View or click the icon in the editor title to run commands.

### Open Excel File

1. Select a macro file in Explorer View
2. Select **Open Excel Macro** from the editor title
3. File opens in Excel

### Load VBA

1. Open the macro file in Excel
2. Select the macro file in VS Code's Explorer View
3. Select **Load VBA from Excel Macro** from the editor title
4. A folder with the same name as the macro file is created
   - Example: `MyMacro.xlsm` → `MyMacro_xlsm` folder

### Edit VBA

Edit VBA files in VS Code.

### Compare VBA

1. Select the macro file in Explorer View
2. Select **Compare VBA with Excel Macro** from the editor title
3. Differences with the macro file are displayed

### Save VBA

1. Select the macro file in Explorer View
2. Select **Save VBA to Excel Macro** from the editor title
3. Changes are saved to the macro file
   - For .xlam files, you cannot save directly from the extension. Save from the VB Editor toolbar.

### Run Sub Procedure

1. Open the VBA file in the editor
2. Place the cursor inside the Sub procedure you want to run
3. Select **Run VBA Sub at Cursor** from the editor title
4. The following processes are executed automatically:
   - Save changes to the macro file
   - Execute the Sub procedure in Excel

### Load CustomUI

1. Select the macro file in Explorer View
2. Select **Load CustomUI from Excel Macro** from the editor title
3. A folder with the same name as the macro file is created
   - Example: `MyAddin.xlam` → `MyAddin_customUI` folder
   - CustomUI XML files (customUI.xml, customUI14.xml) are extracted directly into the folder

### Edit CustomUI

Edit CustomUI XML files (customUI.xml, customUI14.xml) in VS Code.

### Save CustomUI

1. Select the macro file in Explorer View
2. Select **Save CustomUI to Excel Macro** from the editor title
3. Changes are saved to the macro file's CustomUI

## Notes

- Files downloaded from the internet require access block removal.
- VBA file encoding is Shift_JIS.
- The .frx file corresponding to .frm is auto-generated. If you delete .frx, .frm cannot be saved.
- Manage macro files with a version control system as backup in case they become corrupted.

---

## Excel VBA Extension JP

VS Code で Excel VBA を開発するための拡張機能です。

**主な利点：**

- VBA ファイルの VS Code での編集
- VS Code からのマクロの実行
- VBA ファイルのバージョン管理との統合
- GitHub Copilot などの生成 AI アシスタントの活用

![excel-vba](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba.png?raw=true)

## 機能

| 機能                               | 説明                                                                          |
| ---------------------------------- | ----------------------------------------------------------------------------- |
| **Open Excel Macro**               | エクスプローラービューから Excel ファイルを開く。                           |
| **Load VBA from Excel Macro**      | マクロファイル(.xlsm, .xlam)から VBA ファイル（.bas, .cls, .frm）を取り出す。 |
| **Save VBA to Excel Macro**        | VS Code で編集した VBA ファイルをマクロファイルに保存します。                 |
| **Compare VBA with Excel Macro**   | VS Code で編集したコードとマクロファイルの状態を比較する。                    |
| **Run VBA Sub at Cursor**          | カーソル位置の Sub プロシージャを Excel で実行します。                        |
| **Load CustomUI from Excel Macro** | マクロファイルから CustomUI（customUI.xml, customUI14.xml）をエクスポート。   |
| **Save CustomUI to Excel Macro**   | VS Code で編集した CustomUI をマクロファイルに保存。                          |

## 準備

### マクロセキュリティ設定

拡張機能が Excel VBA にアクセスするための設定をします。

1. Excel を開く
2. **ファイル** → **オプション** → **セキュリティセンター（トラストセンター）** を選択
3. **セキュリティセンターの設定（トラストセンターの設定）** をクリック
4. **マクロのセキュリティ** セクションで以下を確認：
   - **VBA プロジェクトオブジェクトモデルへのアクセスを信頼する** にチェックを入れる
5. **OK** をクリックして設定を保存

## 使い方

エクスプローラービューでマクロファイルを右クリック、またはエディタタイトルのアイコンをクリックして実行します。

### Excel ファイルを開く

1. エクスプローラービューでマクロファイルを選択
2. エディタのタイトルから「Open Excel Macro」を選択
3. Excel でファイルが開く

### VBA をロードする

1. Excel で編集対象のマクロファイルを開く
2. VS Code のエクスプローラービューでマクロファイルを選択
3. エディタのタイトルから「Load VBA from Excel Macro」を選択
4. マクロファイルと同じ名前のフォルダが作成される
   - 例：`MyMacro.xlsm` → `MyMacro_xlsm` フォルダ

### VBA を編集する

フォルダ内の VBA ファイルを VS Code で編集。

### 差異を確認する

1. エクスプローラービューでマクロファイルを選択
2. エディタのタイトルから「Compare VBA with Excel Macro」を選択
3. マクロファイルとの差異が表示される

### VBA を保存する

1. エクスプローラービューでマクロファイルを選択
2. エディタのタイトルから「Save VBA to Excel Macro」を選択
3. 編集内容がマクロファイルに保存される
   - .xlam の場合は拡張機能から直接保存することはできません。VBエディタのツールバーから保存してください。

### Sub プロシージャを実行する

1. VBA ファイルをエディタで開く
2. 実行したい Sub プロシージャの中にカーソルを置く
3. エディタのタイトルから「Run VBA Sub at Cursor」を選択
4. 以下の処理が自動的に実行される：
   - マクロファイルに VBA の変更内容を保存
   - 指定した Sub プロシージャを Excel で実行

### CustomUI をロードする

1. マクロファイルをエクスプローラービューで選択
2. エディタのタイトルから「Load CustomUI from Excel Macro」を選択
3. アドインファイルと同じ名前のフォルダが作成される
   - 例：`MyAddin.xlam` → `MyAddin_customUI` フォルダ
   - CustomUI XMLファイル（`customUI.xml`, `customUI14.xml`）が直接フォルダ内に抽出される

### CustomUI を編集する

フォルダ内の CustomUI XMLファイル（customUI.xml, customUI14.xml）を VS Code で編集。

### CustomUI を保存する（.xlam のみ）

1. エクスプローラービューでマクロファイルを選択
2. エディタのタイトルから「Save CustomUI to Excel Macro」を選択
3. 編集内容がアドインファイルの CustomUI に保存される

## 注意事項

- ネットからダウンロードしたファイルはアクセスブロックの解除が必要です。
- VBAファイルの文字エンコーティングは Shift_JIS です。
- .frmに対応する.frxは自動生成されます。.frx を削除すると、.frmは保存できません。
- マクロファイルが壊れた場合に備えてバージョン管理システムで管理してください。
