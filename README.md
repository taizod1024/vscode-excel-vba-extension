[ENGLISH](#excel-vba-extension-en) | [日本語](#excel-vba-extension-jp)

**The folder structure has been changed since version 0.2.0.**

- aaa_bas/ → aaa_xlsx/bas/
- aaa_csv/ → aaa_xslx/csv/
- aaa_png/ → aaa_xlsx/png/
- aaa_xml/ → aaa_xlsx/xml/

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

![excel-vba-extension](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba-extension.png?raw=true)

![excel-vba-addin](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba-addin.png?raw=true)

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
| **Export Sheets as PNG**          | Export sheets with print area as PNG images.                                  |
| **Open Sheet from PNG**           | Open the original Excel sheet from a PNG image file.                          |

## Setup

### Macro Security Settings

Configure settings to allow the extension to access Excel VBA.

1. Open Excel
2. Select **File** → **Options** → **Trust Center**
3. Click **Trust Center Settings**
4. In the **Macro Settings** section, verify:
   - Check **Trust access to the VBA project object model**
5. Click **OK** to save settings

### Enable Add-in

After installing the extension, you need to enable the `excel-vba-addin` in Excel:

1. Open Excel
2. Select **File** → **Options** → **Trust Center**
3. Click **Trust Center Settings**
4. Select **Trusted Add-in Catalogs** (or **Trusted Add-ins**)
5. In the **Developer Tab** section:
   - Go to **Developer** tab in the Ribbon
   - Click **Excel Add-ins**
   - Check the checkbox next to `excel-vba-addin` to enable it
6. Click **OK** to save settings

### Enable Content (If Prompted)

When you open an Excel file (.xlsm, .xlam) after installing the extension, a **Enable Content** button may appear at the top of the window. If prompted, click the **Enable Content** button to allow the add-in to function properly.

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
   - Example: `MyBook.xlsm` → `MyBook_xlsm/bas` folder

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
   - Example: `MyAddin.xlam` → `MyAddin_xlam/xml` folder
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

### Export Sheet as PNG

1. In Excel, create sheets with names ending in `.png` (e.g., `Dashboard.png`)
2. Set the print area for each sheet (the content you want to export)
3. Hide gridlines and page breaks in the sheet if you don't want them in the image
4. Select the book file in VS Code's Explorer View
5. Select **Export Sheet as PNG** from the editor title
6. A folder with the same name as the book file is created
   - Example: `MyBook.xlsm` → `MyBook_xlsm/png` folder
   - PNG images are created for all sheets ending with `.png`
   - File names match the sheet names (e.g., `Dashboard.png`)

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
  - Excel シートのPNGとしてエクスポート
  - CustomUI XML ファイルの編集、更新
  - VBA ファイルの Sub プロシージャの VS Code からの実行
- VS Codeとの統合
  - VBA ファイル、Excelシート、CustomUI XMLファイルのバージョン管理
  - VBA ファイル、Excelシート、CustomUI XMLファイルへの生成AIの適用

![excel-vba-extension](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba-extension.png?raw=true)

![excel-vba-addin](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba-addin.png?raw=true)

## 機能

| 機能                                      | 説明                                                                                               |
| ----------------------------------------- | -------------------------------------------------------------------------------------------------- |
| **New Excel Book**                        | 新しい Excel ファイルを作成します。                                                                |
| **New Excel Book with CustomUI as Macro** | CustomUI テンプレート付きの新しい Excel マクロファイルを作成します。                               |
| **New Excel Book with CustomUI as Addin** | CustomUI テンプレート付きの新しい Excel アドインファイルを作成します。                             |
| **Open Excel Book**                       | 選択中のExcel ファイルを開きます。                                                                 |
| **Load VBA from Excel Book**              | ブックファイル(.xlsm, .xlam)から VBA ファイル(.bas, .cls, .frm)を保存します。                      |
| **Save VBA to Excel Book**                | VS Code で編集した VBA ファイルをブックファイルに保存します。                                      |
| **Compare VBA with Excel Book**           | VS Code で編集したコードとブックファイルの状態を比較します。                                       |
| **Run VBA Sub at Cursor**                 | カーソル位置の Sub プロシージャを Excel で実行します。                                             |
| **Load CSV from Excel Book**              | ブックファイルのシート(.csv)を CSV ファイルに保存します。                                          |
| **Save CSV to Excel Book**                | VS Code で編集した CSV ファイルをブックファイルに保存します。                                      |
| **Create URL Shortcut**                   | OneDrive/SharePoint にあるクラウドホストの Excel ファイルのダミー URL ショートカットを作成します。 |
| **Load CustomUI from Excel Book**         | ブックファイルから CustomUI（customUI.xml, customUI14.xml）をXMLファイルに保存します。             |
| **Save CustomUI to Excel Book**           | VS Code で編集した CustomUIのXMLファイルをブックファイルに保存します。                             |
| **Export Sheets as PNG**                  | シートの印刷範囲を PNG 画像ファイルにエクスポートします。                                          |
| **Open Sheet from PNG**                   | PNG 画像ファイルから元の Excel シートを開きます。                                                  |

## 準備

### マクロセキュリティ設定

拡張機能が Excel VBA にアクセスするための設定をします。

1. Excel を開きます。
2. **ファイル** → **オプション** → **セキュリティセンター（トラストセンター）** を選択します。
3. **セキュリティセンターの設定（トラストセンターの設定）** をクリックします。
4. **マクロのセキュリティ** セクションで以下を確認します。
   - **VBA プロジェクトオブジェクトモデルへのアクセスを信頼する** にチェックを入れます。
5. **OK** をクリックして設定を保存します。

### アドインを有効化する

拡張機能をインストール後、Excel で `excel-vba-addin` を有効化する必要があります。

1. Excel を開きます。
2. **ファイル** → **オプション** → **セキュリティセンター（トラストセンター）** を選択します。
3. **セキュリティセンターの設定（トラストセンターの設定）** をクリックします。
4. **信頼できるアドイン カタログ**（または **信頼できるアドイン**）を選択します。
5. **開発タブ** セクションで以下を確認します。
   - リボンの **開発** タブをクリックします。
   - **Excel アドイン** をクリックします。
   - `excel-vba-addin` の横のチェックボックスをチェックして有効化します。
6. **OK** をクリックして設定を保存します。

### コンテンツの有効化（ダイアログが表示された場合）

拡張機能をインストール後、Excel ファイル（.xlsm、.xlam）を開いたときに、ウィンドウの上部に**コンテンツの有効化**ボタンが表示される場合があります。表示された場合は、**コンテンツの有効化** ボタンをクリックして、アドインが正しく機能するようにしてください。

## 使い方

エクスプローラービューでブックファイルを右クリック、またはエディタタイトルのアイコンをクリックして実行します。

### 新しい Excel ブックを作成する

1. メニューから **ファイル** → **新しいファイル** を選択します。
2. 選択肢から **New Excel Book** を選択します。
3. ファイル名を入力（日本語など各国語に対応）します。
   - 注：`.xlsx` 拡張子を除いた名前を入力してください。
   - 例：`MyBook` と入力すると `MyBook.xlsx` として作成されます。
4. 新しい Excel ファイル（.xlsx）が作成されて Excel で開きます。

### CustomUI テンプレート付きの新規 Excel ファイルを作成する

1. コマンドパレット（Ctrl+Shift+P / Cmd+Shift+P）を開きます。
2. **"New Excel Book with CustomUI"** を検索して選択します。
3. ファイル名を入力（日本語など各国語に対応）します。
   - 注：`.xlsm` 拡張子を除いた名前を入力してください。
   - 例：`MyMacro` と入力すると `MyMacro.xlsm` として作成されます。
4. CustomUI テンプレート付きの新しい Excel ファイル（.xlsm）が作成されて開きます。

### Excel ファイルを開く

1. エクスプローラービューでブックファイルを選択します。
2. エディタのタイトルから **Open Excel Book** を選択します。
3. Excel でファイルが開きます。

### VBA をロードする

1. Excel で編集対象のブックファイルを開きます。
2. VS Code のエクスプローラービューでブックファイルを選択します。
3. エディタのタイトルから **Load VBA from Excel Book** を選択します。
4. ブックファイルと同じ名前のフォルダが作成されます。
   - 例：`MyBook.xlsm` → `MyBook_xlsm/bas` フォルダ

### VBA を編集する

フォルダ内の VBA ファイルを VS Code で編集します。

### 差異を確認する

1. エクスプローラービューでブックファイルを選択します。
2. エディタのタイトルから **Compare VBA with Excel Book** を選択します。
3. ブックファイルとの差異が表示されます。

### VBA を保存する

1. エクスプローラービューでブックファイルを選択します。
2. エディタのタイトルから **Save VBA to Excel Book** を選択します。
3. 編集内容がブックファイルに保存されます。

### Sub プロシージャを実行する

1. VBA ファイルをエディタで開きます。
2. 実行したい Sub プロシージャの中にカーソルを置きます。
3. エディタのタイトルから **Run VBA Sub at Cursor** を選択します。
4. 以下の処理が自動的に実行されます。
   - ブックファイルに VBA の変更内容を保存します。
   - 指定した Sub プロシージャを Excel で実行します。

### CustomUI をロードする

1. ブックファイルをエクスプローラービューで選択します。
2. エディタのタイトルから **Load CustomUI from Excel Book** を選択します。
3. アドインファイルと同じ名前のフォルダが作成されます。
   - 例：`MyAddin.xlam` → `MyAddin_xlam/xml` フォルダ
   - CustomUI XMLファイル（`customUI.xml`, `customUI14.xml`）が直接フォルダ内に抽出されます。

### CustomUI を編集する

フォルダ内の CustomUI XMLファイル（customUI.xml, customUI14.xml）を VS Code で編集します。

### CustomUI を保存する（.xlam のみ）

1. エクスプローラービューでブックファイルを選択します。
2. エディタのタイトルから **Save CustomUI to Excel Book** を選択します。
3. 編集内容がアドインファイルの CustomUI に保存されます。

### ダミー URL ショートカットを作成する

OneDrive または SharePoint に保存されている Excel ファイルを処理できるようにします。

1. 処理したいクラウドホストのブックファイルを Excel で開きます（複数可）
2. コマンドパレット（Ctrl+Shift+P / Cmd+Shift+P）を開きます。
3. **Create URL Shortcut** を検索して実行します。
4. ワークスペースフォルダにダミー URL ショートカットファイル（.url）が作成されます。
   - これはクラウドホストの Excel ファイルのマーカーとして機能します。

**仕組み：**

- ダミー URL ファイルはクラウドホストの Excel ファイルのマーカーとして機能します。
- VBA、シート、CustomUI の Load/Save を実行する際、拡張機能はこのマーカーを自動的に検出します。
- 処理時は Excel で開いているアクティブなワークブックが使用されます。（ローカルファイルのコピーは不要）
- これにより、クラウドホストファイルのシームレスなバージョン管理と編集が可能になります。

### シートを PNG 画像にエクスポートする（複数可）

1. Excel で、`.png` で終わる名前のシートを作成します。（例：`Dashboard.png`）
2. 各シートにエクスポートしたい範囲として印刷範囲を設定します。
3. や画像にグリッド線やページ番号を含めたくない場合は、シートで非表示に設定します。
4. VS Code のエクスプローラービューでブックファイルを選択しまうｓ。
5. エディタのタイトルから **Export Sheets as PNG** を選択します。
6. ブックファイルと同じ名前のフォルダが作成されます。
   - 例 : `MyBook.xlsm` → `MyBook_xlsm/png` フォルダ
   - `.png` で終わるすべてのシートの PNG 画像が作成されます。
   - ファイル名はシート名と同じです（例：`Dashboard.png`）

### PNG から Sheet を開く

1. VS Code のエクスプローラービューで PNG ファイルを選択します。（Export Sheet as PNG で作成したファイル）
2. 右クリックして **Open Sheet from PNG** を選択します。
3. 元の Excel ブックが開かれて、PNG ファイル名に対応するシートが自動的に選択されます。
   - 例：`Dashboard.png` ファイル → Excel の `Dashboard.png` というシートが選択されます。
4. PNG をクリックするだけで、元のデータの編集が可能になります。

- **全般**
  - ネットからダウンロードしたファイルはアクセスブロックの解除が必要です。
  - 適切に動作するにはコンテンツの有効化が必要です。
  - ブックファイルが壊れた場合に備えてバージョン管理システムで管理してください。
- **アドインファイル**
  - .xlam の場合は拡張機能から直接保存することはできません。VBエディタのツールバーから保存してください。
- **VBAファイル**
  - VBAファイルの文字エンコーティングは Shift_JIS です。
  - .frmに対応する.frxは自動生成されます。.frx を削除すると、.frmは保存できません。
- **CSVファイル**
  - 名前に".csv"が付いたシートを処理対象とします。
- **PNG ファイル**
  - 名前に".png"が付いたシートを処理対象とします。
- **追加**
  - VBAファイルはVS Codeで新規追加することができます。
  - CSVファイルはVS Codeで新規追加することができます。
  - CustomUI XMLファイルは新規追加することはできません。事前にブックファイルに登録されている必要があります。
