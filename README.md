[ENGLISH](#excel-vba-extension-en) | [日本語](#excel-vba-extension-jp)

---

## Excel VBA Extension EN

A VS Code extension for developing Excel VBA / CSV / CustomUI.

**Key Benefits:**

- Editing in VS Code
  - Edit, add, update, and delete VBA files
  - Edit, add, update, and delete Excel CSV
  - Edit and update CustomUI XML files
  - Execute Sub procedures of VBA files from VS Code
- Integration with VS Code
  - Version control for VBA files, Excel CSV, and CustomUI XML files
  - Apply generative AI to VBA files, Excel CSV, and CustomUI XML files

![excel-vba](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba.png?raw=true)

## Features

| Feature                           | Description                                                                   |
| --------------------------------- | ----------------------------------------------------------------------------- |
| **New Excel Book**                | Create a new Excel file.                                                      |
| **New Excel Book with CustomUI**  | Create a new Excel file with CustomUI template.                               |
| **Open Excel Book**               | Open Excel files from Explorer View.                                          |
| **Load VBA from Excel Book**      | Extract VBA files (.bas, .cls, .frm) from book files (.xlsm, .xlam).          |
| **Save VBA to Excel Book**        | Save VBA files edited in VS Code to book files.                               |
| **Compare VBA with Excel Book**   | Compare code edited in VS Code with the state of book files.                  |
| **Run VBA Sub at Cursor**         | Execute the Sub procedure at cursor position in Excel.                        |
| **Load CSV from Excel Book**      | Export sheets (.csv) from book files as CSV files.                            |
| **Save CSV to Excel Book**        | Save CSV files edited in VS Code to book file sheets.                         |
| **Create URL Shortcut**           | Create URL shortcut files for cloud-hosted Excel files (OneDrive/SharePoint). |
| **Load CustomUI from Excel Book** | Export CustomUI (customUI.xml, customUI14.xml) from book files.               |
| **Save CustomUI to Excel Book**   | Save CustomUI edited in VS Code to book files.                                |

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

Right-click a book file in Explorer View or click the icon in the editor title to run commands.

### New Excel Book

1. **File** → **New File** from the menu
2. Select **New Excel Book** from the options
3. Enter the file name (supports Japanese and other characters)
   - Note: Enter the name without the `.xlsx` extension
   - Example: `MyBook` (the file will be created as `MyBook.xlsx`)
4. A new Excel file (.xlsx) is created and opened in Excel

### New Excel Book with CustomUI

1. Open the Command Palette (Ctrl+Shift+P / Cmd+Shift+P)
2. Search for and select **"New Excel Book with CustomUI"**
3. Enter the file name (supports Japanese and other characters)
   - Note: Enter the name without the `.xlsm` extension
   - Example: `MyMacro` (the file will be created as `MyMacro.xlsm`)
4. A new Excel file (.xlsm) with CustomUI template is created and opened

### Open Excel File

1. Select a book file in Explorer View
2. Select **Open Excel Book** from the editor title
3. File opens in Excel

### Load VBA

1. Open the book file in Excel
2. Select the book file in VS Code's Explorer View
3. Select **Load VBA from Excel Book** from the editor title
4. A folder with the same name as the book file is created
   - Example: `MyBook.xlsm` → `MyBook_bas` folder

### Edit VBA

Edit VBA files in VS Code.

### Compare VBA

1. Select the book file in Explorer View
2. Select **Compare VBA with Excel Book** from the editor title
3. Differences with the book file are displayed

### Save VBA

1. Select the book file in Explorer View
2. Select **Save VBA to Excel Book** from the editor title
3. Changes are saved to the book file
   - For .xlam files, you cannot save directly from the extension. Save from the VB Editor toolbar.

### Run Sub Procedure

1. Open the VBA file in the editor
2. Place the cursor inside the Sub procedure you want to run
3. Select **Run VBA Sub at Cursor** from the editor title
4. The following processes are executed automatically:
   - Save changes to the book file
   - Execute the Sub procedure in Excel

### Load CustomUI

1. Select the book file in Explorer View
2. Select **Load CustomUI from Excel Book** from the editor title
3. A folder with the same name as the book file is created
   - Example: `MyAddin.xlam` → `MyAddin_xml` folder
   - CustomUI XML files (customUI.xml, customUI14.xml) are extracted directly into the folder

### Edit CustomUI

Edit CustomUI XML files (customUI.xml, customUI14.xml) in VS Code.

### Save CustomUI

1. Select the book file in Explorer View
2. Select **Save CustomUI to Excel Book** from the editor title
3. Changes are saved to the book file's CustomUI

### Create URL Shortcut

For Excel files stored on OneDrive or SharePoint:

1. Open all the cloud-hosted book files in Excel that you want to process
2. Open the Command Palette (Ctrl+Shift+P / Cmd+Shift+P)
3. Search for and select **"Create URL Shortcut"**
4. URL shortcut files (.url) are created in the workspace folder
   - These act as markers for cloud-hosted Excel files

**How it works:**

- URL files serve as markers for cloud-hosted Excel files
- When you Load/Save VBA, Sheets, or CustomUI, the extension automatically detects these markers
- The active Excel workbook is used for processing (no local file copy needed)
- This enables seamless version control and editing of cloud-hosted files

## Notes

- **General**
  - Files downloaded from the internet require unblocking access.
  - Manage book files with a version control system in case they become corrupted.
- **Add-in Files**
  - For .xlam files, they cannot be saved directly from the extension. Please save from the VB Editor toolbar.
- **VBA Files**
  - VBA files use Shift_JIS character encoding.
  - .frx corresponding to .frm is automatically generated. If you delete .frx, .frm cannot be saved.
- **CSV Files**
  - Sheets with ".csv" in the name are treated as targets for processing.
- **Additional**
  - VBA files can be created new in VS Code.
  - CSV files can be created new in VS Code.
  - CustomUI XML files cannot be created new. They must be registered in the book file beforehand.

---

## Excel VBA Extension JP

VS Code で Excel VBA を開発するための拡張機能です。

**主な利点：**

- VS Codeでの編集
  - VBA ファイルの編集、追加・更新・削除
  - Excel シートの編集、追加・更新・削除
  - CustomUI XML ファイルの編集、更新
  - VBA ファイルの Sub プロシージャの VS Code からの実行
- VS Codeとの統合
  - VBA ファイル、Excelシート、CustomUI XMLファイルのバージョン管理
  - VBA ファイル、Excelシート、CustomUI XMLファイルへの生成AIの適用

![excel-vba](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba.png?raw=true)

## 機能

| 機能                              | 説明                                                                                             |
| --------------------------------- | ------------------------------------------------------------------------------------------------ |
| **New Excel Book**                | 新しい Excel ファイルを作成する。                                                                |
| **New Excel Book with CustomUI**  | CustomUI テンプレート付きの新しい Excel ファイルを作成する。                                     |
| **Open Excel Book**               | エクスプローラービューから Excel ファイルを開く。                                                |
| **Load VBA from Excel Book**      | ブックファイル(.xlsm, .xlam)から VBA ファイル(.bas, .cls, .frm)を読み込む。                      |
| **Save VBA to Excel Book**        | VS Code で編集した VBA ファイルをブックファイルに保存する。                                      |
| **Compare VBA with Excel Book**   | VS Code で編集したコードとブックファイルの状態を比較する。                                       |
| **Run VBA Sub at Cursor**         | カーソル位置の Sub プロシージャを Excel で実行する。                                             |
| **Load CSV from Excel Book**      | ブックファイルのシート(.csv)を CSV ファイルに保存する。                                          |
| **Save CSV to Excel Book**        | VS Code で編集した CSV ファイルをブックファイルに保存する。                                      |
| **Create URL Shortcut**           | OneDrive/SharePoint にあるクラウドホストの Excel ファイルのダミー URL ショートカットを作成する。 |
| **Load CustomUI from Excel Book** | ブックファイルから CustomUI（customUI.xml, customUI14.xml）をエクスポートする。                  |
| **Save CustomUI to Excel Book**   | VS Code で編集した CustomUI をブックファイルに保存する。                                         |

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

エクスプローラービューでブックファイルを右クリック、またはエディタタイトルのアイコンをクリックして実行します。

### 新しい Excel ブックを作成する

1. メニューから **ファイル** → **新しいファイル** を選択
2. 選択肢から **New Excel Book** を選択
3. ファイル名を入力（日本語など各国語に対応）
   - 注：`.xlsx` 拡張子を除いた名前を入力してください
   - 例：`MyBook` と入力すると `MyBook.xlsx` として作成されます
4. 新しい Excel ファイル（.xlsx）が作成されて Excel で開く

### CustomUI テンプレート付きの新規 Excel ファイルを作成する

1. コマンドパレット（Ctrl+Shift+P / Cmd+Shift+P）を開く
2. **"New Excel Book with CustomUI"** を検索して選択
3. ファイル名を入力（日本語など各国語に対応）
   - 注：`.xlsm` 拡張子を除いた名前を入力してください
   - 例：`MyMacro` と入力すると `MyMacro.xlsm` として作成されます
4. CustomUI テンプレート付きの新しい Excel ファイル（.xlsm）が作成されて開く

### Excel ファイルを開く

1. エクスプローラービューでブックファイルを選択
2. エディタのタイトルから「Open Excel Book」を選択
3. Excel でファイルが開く

### VBA をロードする

1. Excel で編集対象のブックファイルを開く
2. VS Code のエクスプローラービューでブックファイルを選択
3. エディタのタイトルから「Load VBA from Excel Book」を選択
4. ブックファイルと同じ名前のフォルダが作成される
   - 例：`MyBook.xlsm` → `MyBook_bas` フォルダ

### VBA を編集する

フォルダ内の VBA ファイルを VS Code で編集。

### 差異を確認する

1. エクスプローラービューでブックファイルを選択
2. エディタのタイトルから「Compare VBA with Excel Book」を選択
3. ブックファイルとの差異が表示される

### VBA を保存する

1. エクスプローラービューでブックファイルを選択
2. エディタのタイトルから「Save VBA to Excel Book」を選択
3. 編集内容がブックファイルに保存される

### Sub プロシージャを実行する

1. VBA ファイルをエディタで開く
2. 実行したい Sub プロシージャの中にカーソルを置く
3. エディタのタイトルから「Run VBA Sub at Cursor」を選択
4. 以下の処理が自動的に実行される：
   - ブックファイルに VBA の変更内容を保存
   - 指定した Sub プロシージャを Excel で実行

### CustomUI をロードする

1. ブックファイルをエクスプローラービューで選択
2. エディタのタイトルから「Load CustomUI from Excel Book」を選択
3. アドインファイルと同じ名前のフォルダが作成される
   - 例：`MyAddin.xlam` → `MyAddin_xml` フォルダ
   - CustomUI XMLファイル（`customUI.xml`, `customUI14.xml`）が直接フォルダ内に抽出される

### CustomUI を編集する

フォルダ内の CustomUI XMLファイル（customUI.xml, customUI14.xml）を VS Code で編集。

### CustomUI を保存する（.xlam のみ）

1. エクスプローラービューでブックファイルを選択
2. エディタのタイトルから「Save CustomUI to Excel Book」を選択
3. 編集内容がアドインファイルの CustomUI に保存される

### ダミー URL ショートカットを作成する

OneDrive または SharePoint に保存されている Excel ファイルを処理する場合：

1. 処理したいクラウドホストのブックファイルを Excel で開く（複数可）
2. コマンドパレット（Ctrl+Shift+P / Cmd+Shift+P）を開く
3. **「Create URL Shortcut」** を検索して実行
4. ワークスペースフォルダにダミー URL ショートカットファイル（.url）が作成される
   - これはクラウドホストの Excel ファイルのマーカーとして機能する

**仕組み：**

- ダミー URL ファイルはクラウドホストの Excel ファイルのマーカーとして機能する
- VBA、シート、CustomUI の Load/Save を実行する際、拡張機能はこのマーカーを自動的に検出する
- 処理時は Excel で開いているアクティブなワークブックが使用される（ローカルファイルのコピーは不要）
- これにより、クラウドホストファイルのシームレスなバージョン管理と編集が可能になる

## 注意事項

- **全般**
  - ネットからダウンロードしたファイルはアクセスブロックの解除が必要です。
  - ブックファイルが壊れた場合に備えてバージョン管理システムで管理してください。
- **アドインファイル**
  - .xlam の場合は拡張機能から直接保存することはできません。VBエディタのツールバーから保存してください。
- **VBAファイル**
  - VBAファイルの文字エンコーティングは Shift_JIS です。
  - .frmに対応する.frxは自動生成されます。.frx を削除すると、.frmは保存できません。
- **CSVファイル**
  - 名前に".csv"が付いたシートを処理対象とします。
- **追加**
  - VBAファイルはVS Codeで新規追加することができます。
  - CSVファイルはVS Codeで新規追加することができます。
  - CustomUI XMLファイルは新規追加することはできません。事前にブックファイルに登録されている必要があります。
