## Excel VBA

![excel-vba](https://github.com/taizod1024/vscode-excel-vba-extension/blob/main/images/excel-vba.png?raw=true)

VS Code で Excel VBA を開発するための拡張機能です。

**主な利点：**

- VBAファイルのVS Codeでの編集
- VBAファイルのバージョン管理との統合
- GitHub Copilot などの生成AI アシスタントの活用
- VS Codeからのマクロの実行

## 機能

| 機能                                | 説明                                                                                   |
| ----------------------------------- | -------------------------------------------------------------------------------------- |
| **Open Excel Macro**                | ファイルエクスプローラーから Excel ファイルを開きます。                                |
| **Load VBA from Excel Macro**       | マクロファイルから VBA ファイル（.bas, .cls, .frm）を取り出す。                        |
| **Save VBA to Excel Macro**         | VS Code で編集した VBA ファイルをマクロファイルに保存します。                          |
| **Compare VBA with Excel Macro**    | VS Code で編集したコードとマクロファイルの状態を比較する。                             |
| **Run VBA Sub at Cursor**           | カーソル位置の Sub プロシージャを Excel で実行します。                                  |
| **Load CustomUI from Excel Add-in** | アドインファイル（.xlam）から CustomUI（customUI.xml, customUI14.xml）をエクスポート。 |
| **Save CustomUI to Excel Add-in**   | VS Code で編集した CustomUI をアドインファイル（.xlam）に保存。                        |

## 準備

### マクロセキュリティ設定

拡張機能が Excel の VBA モジュールにアクセスするための設定をします。

1. Excel を開く
2. **ファイル** → **オプション** → **セキュリティセンター（トラストセンター）** を選択
3. **セキュリティセンターの設定（トラストセンターの設定）** をクリック
4. **マクロのセキュリティ** セクションで以下を確認：
   - **VBA プロジェクトオブジェクトモデルへのアクセスを信頼する** にチェックを入れる
5. **OK** をクリックして設定を保存

## 使い方

ファイルエクスプローラーでマクロファイルを右クリック、またはエディタタイトルのアイコンをクリックして実行します。

### Excel ファイルを開く

1. ファイルエクスプローラーでマクロファイルを右クリック
2. 「Open Excel Macro」を選択
3. Excel でファイルが開く

### VBA をロードする

1. Excel で編集対象のマクロファイルを開く
2. VS Code のファイルエクスプローラーでマクロファイルを右クリック
3. 「Load VBA from Excel Macro」を選択
4. マクロファイルと同じ名前のフォルダが作成される
   - 例：`MyMacro.xlsm` → `MyMacro_xlsm` フォルダ

### VBA を編集する

フォルダ内の VBA ファイル（.bas, .cls, .frm）を VS Code で編集。

### 差異を確認する

1. ファイルエクスプローラーでマクロファイルを右クリック
2. 「Compare VBA with Excel Macro」を選択
3. マクロファイルとの差異が表示される

### VBA を保存する

1. ファイルエクスプローラーでマクロファイルを右クリック
2. 「Save VBA to Excel Macro」を選択
3. 編集内容がマクロファイルに保存される
   - .xlam の場合は拡張機能から直接保存することはできません。VBエディタのツールバーから保存してください。

### Sub プロシージャを実行する

1. VBA ファイル（.bas, .cls, .frm）をエディタで開く
2. 実行したい Sub プロシージャの中にカーソルを置く
3. ファイルエクスプローラーでマクロファイルを右クリック（または エディタタイトルのアイコン）
4. 「Run VBA Sub at Cursor」を選択
5. 以下の処理が自動的に実行される：
   - マクロファイルに VBA の変更内容を保存
   - 指定した Sub プロシージャを Excel で実行

### CustomUI をロードする（.xlam のみ）

1. アドインファイル（.xlam）をファイルエクスプローラーで右クリック
2. 「Load CustomUI from Excel Add-in」を選択
3. アドインファイルと同じ名前のフォルダが作成される
   - 例：`MyAddin.xlam` → `MyAddin_customUI` フォルダ
   - CustomUI XMLファイル（`customUI.xml`, `customUI14.xml`）が直接フォルダ内に抽出される

### CustomUI を編集する

フォルダ内の CustomUI XMLファイル（customUI.xml, customUI14.xml）を VS Code で編集。

### CustomUI を保存する（.xlam のみ）

1. ファイルエクスプローラーでアドインファイル（.xlam）を右クリック
2. 「Save CustomUI to Excel Add-in」を選択
3. 編集内容がアドインファイルの CustomUI に保存される

## 注意事項

- ネットからダウンロードしたファイルはアクセスブロックの解除が必要です。
- .bas, .cls, .frmの文字エンコーティングは Shift_JIS です。
- .frmに対応する.frxは自動生成されます。.frx を削除すると、.frmは保存できません。
- .frmを保存する際にコードの先頭に空行が追加されます。 https://learn.microsoft.com/en-us/answers/questions/4911760/excel-vba-bug-importing-a-form-adds-a-newline-at-t
- Sub を実行するには、Excel が起動している必要があります。
- Sub を実行する際に VBA プロジェクトオブジェクトモデルへのアクセス権が必要です。詳細は「準備」セクションを参照してください。
- CustomUIは .xlamでのみサポートします。
- Excel マクロが壊れた場合に備えてバージョン管理システムで管理してください。
