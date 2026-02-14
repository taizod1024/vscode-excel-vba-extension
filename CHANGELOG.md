# Changelog

All notable changes to this project will be documented in this file. See [standard-version](https://github.com/conventional-changelog/standard-version) for commit guidelines.

### [0.2.2](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.2.1...v0.2.2) (2026-02-14)


### Bug Fixes

* READMEとCustomUIのラベルを更新し、Excel VBAの表記を修正 ([cec14d8](https://github.com/taizod1024/vscode-excel-vba-extension/commit/cec14d8c7735e01ac7ad9f24de49f2c0c884fab1))

### [0.2.1](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.2.0...v0.2.1) (2026-02-14)


### Features

* Export-SheetAsImage.ps1をExport-SheetAsPng.ps1にリネームし、PNGエクスポート機能を実装 ([38d6dca](https://github.com/taizod1024/vscode-excel-vba-extension/commit/38d6dca1465e80718e26b171587993a49b2bf2f1))
* Export-SheetAsPngコマンドのIDをexportSheetsAsPngに変更し、関連する関数名を更新 ([1d0f95a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1d0f95a863dce203b9c567e755ec7236ccf5bce4))
* PNGエクスポート機能を追加し、PNGからシートを開く機能を実装 ([d3dd20e](https://github.com/taizod1024/vscode-excel-vba-extension/commit/d3dd20e3b4ff56bd7780cb4322f94d4b7cbd9cfb))
* PNGからシートを開く機能を追加 ([fe494fa](https://github.com/taizod1024/vscode-excel-vba-extension/commit/fe494fa11fae5b6d1a6a9b4fab60d41ec10d9f9e))


### Bug Fixes

* Excelの一時ファイルを正しく無視するために.gitignoreを更新 ([b199880](https://github.com/taizod1024/vscode-excel-vba-extension/commit/b199880a32f0bbfdd00c6fd84b741970d1356e48))
* ExportSheetsAsPngサブルーチンのエラーメッセージを修正し、PowerShellスクリプトの存在確認を追加 ([fce0b76](https://github.com/taizod1024/vscode-excel-vba-extension/commit/fce0b760e4e3cd0630bf428c4f52ac2614ec56a1))
* PNGファイル選択時の説明文を修正 ([92f6507](https://github.com/taizod1024/vscode-excel-vba-extension/commit/92f6507f5658eb988b22a176bbaca83e7da98b84))
* READMEを更新し、ExcelシートのPNGエクスポート機能を追加 ([572a77c](https://github.com/taizod1024/vscode-excel-vba-extension/commit/572a77c208499c59c0acf704bd69ee8c8916053d))
* READMEを更新し、重要な注意事項を追加 ([4caed37](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4caed373058e16e5447cef81107b4d964f46c7e7))
* SampleMacroサブルーチンの構文を修正し、コマンド実行時に出力チャネルを表示するように更新 ([50724c5](https://github.com/taizod1024/vscode-excel-vba-extension/commit/50724c58f94cc8034011eea9e1d99c907a793b1b))
* コマンドの出力メッセージを改善し、URLショートカット作成時の詳細を追加 ([8f65318](https://github.com/taizod1024/vscode-excel-vba-extension/commit/8f653181fa34d7ec254830c88d9524defd3a5e02))
* コマンドの条件をresourcePathに基づいて更新 ([1865641](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1865641ba9d20fcfa993480862d280f0e0af0654))
* 修正されたエラーハンドリングの構文を改善 ([71fe7fa](https://github.com/taizod1024/vscode-excel-vba-extension/commit/71fe7fa3ca6d77e5cdab89e69522734e4756f468))
* 不要なエクスポート機能のボタンをカスタムUIから削除 ([2512d1c](https://github.com/taizod1024/vscode-excel-vba-extension/commit/2512d1c78aad2201468df8d195573f5aefa386a2))

## [0.2.0](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.23...v0.2.0) (2026-02-13)


### ⚠ BREAKING CHANGES

* **FOLDER NAME CHANGED** .xlsx.bas, .xlsx.csv, .xlsx.xml, .xlsx.png

### Features

* **FOLDER NAME CHANGED** .xlsx.bas, .xlsx.csv, .xlsx.xml, .xlsx.png ([6d77592](https://github.com/taizod1024/vscode-excel-vba-extension/commit/6d77592ca4096ad5b5ed91df856ffdf8aa9598ec))
* Excelシートを画像としてエクスポートする機能を追加 ([0ff15c3](https://github.com/taizod1024/vscode-excel-vba-extension/commit/0ff15c3df7c778b4d22d56c7c650d91c25c5de4d))
* Excelファイル名を解決し、進行状況通知のタイトルを更新 ([192be0a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/192be0a3d80796882c044d11c35fe8a8555d9e5b))
* VBA関連コマンドのロギング機能を追加し、エラーメッセージを改善 ([e7fc8af](https://github.com/taizod1024/vscode-excel-vba-extension/commit/e7fc8af0f90cb619ad6ba9f1025ca5e42868c768))
* シートをPNG画像としてエクスポートする機能を追加 ([2bc368e](https://github.com/taizod1024/vscode-excel-vba-extension/commit/2bc368e32b766f8904f262036e95810ba99bfafd))
* 新しいカスタムUIとサンプルモジュールを追加し、エラーメッセージを改善 ([0d03b68](https://github.com/taizod1024/vscode-excel-vba-extension/commit/0d03b680355e3063169d37654a6e1f18288b8a4f))
* 新機能追加 - Excel シートを PNG 画像としてエクスポートする機能仕様書を作成 ([902c030](https://github.com/taizod1024/vscode-excel-vba-extension/commit/902c030b4c5c57d923e72b6f4abd62d146df262c))


### Bug Fixes

* COM呼び出しの信頼性を向上させるために、リストオブジェクトの追加方法を修正 ([1c1d1d5](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1c1d1d5a291aeb73252ec3dcf66ed371b3eaa6b5))
* エラーメッセージを改善し、ユーザーに表示する内容を明確化 ([ad4321a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/ad4321aee77fea037e9cfff73161f71a428822c9))

### [0.1.23](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.22...v0.1.23) (2026-02-12)

### [0.1.22](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.21...v0.1.22) (2026-02-12)

### [0.1.21](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.20...v0.1.21) (2026-02-12)


### Bug Fixes

* simplify error message in Find-VBProject function for clarity ([2087790](https://github.com/taizod1024/vscode-excel-vba-extension/commit/2087790edae7425d394b4833ee3c52fb859942a6))

### [0.1.20](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.19...v0.1.20) (2026-02-12)

### [0.1.19](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.18...v0.1.19) (2026-02-12)


### Features

* add file encoding configuration for URL files ([45e925f](https://github.com/taizod1024/vscode-excel-vba-extension/commit/45e925f19d4edb99ac36e25c54e9cdfebe67ceb1))
* add iconv-lite for SJIS encoding support in URL file handling ([5d24323](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5d243235a27765349366c1744323949f72568e9c))


### Bug Fixes

* update error messages for clarity in GetExtensionPath function and add URL file association ([5dd4558](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5dd45580365eccbfad230392ea80e78bfc1f5b10))

### [0.1.18](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.17...v0.1.18) (2026-02-12)


### Features

* add new command for creating a workbook with custom UI ([590098e](https://github.com/taizod1024/vscode-excel-vba-extension/commit/590098ebd063bb0b1940e9eb72d3a2ae5db815da))
* enhance OpenVSCode functionality to handle recent files and URL paths ([faf7130](https://github.com/taizod1024/vscode-excel-vba-extension/commit/faf7130abeb9003f8f24ced1adb92d0b73aa654e))


### Bug Fixes

* update error messages in OpenVSCode for clarity ([68d8c8a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/68d8c8a18d29dc9a85c7d9fb1e478c9acb6dec81))

### [0.1.17](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.16...v0.1.17) (2026-02-11)


### Bug Fixes

* change VS Code launch command to run in hidden mode ([e1cf568](https://github.com/taizod1024/vscode-excel-vba-extension/commit/e1cf56894ff2402601d9f334cb395679f7a4fa16))

### [0.1.16](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.15...v0.1.16) (2026-02-11)


### Features

* add instructions to enable Excel add-in and content for proper functionality ([6b16f25](https://github.com/taizod1024/vscode-excel-vba-extension/commit/6b16f25a70c6508af5035084f9cfca6e36b6a8c4))
* implement uninstall script to clean up Excel add-in files ([ec6f9de](https://github.com/taizod1024/vscode-excel-vba-extension/commit/ec6f9dec0feaa64a9da5cb62b2d459feaf11f510))


### Bug Fixes

* remove preuninstall script for add-in cleanup ([444ba2b](https://github.com/taizod1024/vscode-excel-vba-extension/commit/444ba2bbfb452a43e875c846bc3ced8b06877f28))
* remove unnecessary blank line in activate function ([3359135](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3359135e2ccf5cc37903cb9a7767f8c2b8f46976))

### [0.1.15](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.14...v0.1.15) (2026-02-11)


### Features

* add command to create new Excel book with CustomUI ([35b8ff6](https://github.com/taizod1024/vscode-excel-vba-extension/commit/35b8ff6b561dc48016e4269ca4cf0cc58dac1884))
* Add commands for Excel workbook management and VBA operations ([d6c6cef](https://github.com/taizod1024/vscode-excel-vba-extension/commit/d6c6cef18591be9ac9e976b48a8e84e7d14715ef))
* add documentation for creating new Excel book with CustomUI ([92b0cb2](https://github.com/taizod1024/vscode-excel-vba-extension/commit/92b0cb28bf84a2372cf3e324fffe5534909e21d6))
* add sample ([e7a7e49](https://github.com/taizod1024/vscode-excel-vba-extension/commit/e7a7e49919c91db25121373557da342b796111e1))


### Bug Fixes

* remove unnecessary whitespace in template file path ([0bed303](https://github.com/taizod1024/vscode-excel-vba-extension/commit/0bed303a96becb770fb3f1baee728500dac59363))
* simplify temp directory fallback paths in uninstall script ([e6195a1](https://github.com/taizod1024/vscode-excel-vba-extension/commit/e6195a10845f8bd4a24404b0247bdff123055baf))
* streamline file name prompts and ensure correct extensions for new workbooks and shortcuts ([420a9eb](https://github.com/taizod1024/vscode-excel-vba-extension/commit/420a9eb7f399bc11f74d6aaa0a91dda7ef3a92bb))
* update command groups for Excel VBA navigation commands ([f854e8b](https://github.com/taizod1024/vscode-excel-vba-extension/commit/f854e8b9ef363e98eabf04e1ddd0606d9c8049c8))
* update prompt text for new workbook name in input dialogs ([6718053](https://github.com/taizod1024/vscode-excel-vba-extension/commit/67180537b0586051eb0b0f8cc3aeef9193c3e108))

### [0.1.14](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.13...v0.1.14) (2026-02-11)


### Features

* add new Excel VBA macro file for enhanced ([9daf099](https://github.com/taizod1024/vscode-excel-vba-extension/commit/9daf09989cf97b1445eb3c08f5d50c2068d532bf))


### Bug Fixes

* correct log message and improve file existence check in Load-CustomUI.ps1 ([ccb77b8](https://github.com/taizod1024/vscode-excel-vba-extension/commit/ccb77b890a7f707fe5586aa7fe7d67eaf2387ce3))
* rename macroPath to bookPath for consistency across scripts ([02a6ccd](https://github.com/taizod1024/vscode-excel-vba-extension/commit/02a6ccd264d56f1fad74fed6b582389f0bf1c3ea))
* standardize parameter naming for consistency across scripts ([3bc5939](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3bc5939d9dacecca1a3c826bd7b1280f23837951))
* update log messages for clarity and consistency in Load-CustomUI.ps1, Save-CustomUI.ps1, and Save-VBA.ps1 ([07ccf49](https://github.com/taizod1024/vscode-excel-vba-extension/commit/07ccf49b2c4293e54fb0716934df5142a09c3074))

### [0.1.13](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.12...v0.1.13) (2026-02-09)


### Bug Fixes

* streamline Excel application references and improve freeze pane handling in Save-CSV.ps1 ([1b2eb18](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1b2eb188fe186dcf1e72f43cae6b9d0a0a35018a))

### [0.1.12](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.11...v0.1.12) (2026-02-08)


### Bug Fixes

* correct spelling of 'saved' in Save-VBA.ps1 for consistency ([3a43624](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3a436249b5f11f1d122d4049a9bc53467c9d521a))
* improve comments for clarity and enhance file handling in Save-VBA.ps1 ([06ef0c5](https://github.com/taizod1024/vscode-excel-vba-extension/commit/06ef0c5adc8d16b316e9a6de0c27903d52209d32))
* improve component removal logic and add verification for standard modules in Save-VBA.ps1 ([15ea7e1](https://github.com/taizod1024/vscode-excel-vba-extension/commit/15ea7e1417f867a846e703f35a50453bf746f38d))

### [0.1.11](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.10...v0.1.11) (2026-02-06)


### Bug Fixes

* set Excel range format to text in Update-SheetData function ([c354c33](https://github.com/taizod1024/vscode-excel-vba-extension/commit/c354c33b5051e2e85975fc8f7c7510d0d72156c6))

### [0.1.10](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.9...v0.1.10) (2026-02-06)


### Bug Fixes

* update table style in Save-CSV.ps1 for improved visual consistency ([30bdb56](https://github.com/taizod1024/vscode-excel-vba-extension/commit/30bdb560638b5e9da6d39c81da663781c495cecd))

### [0.1.9](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.8...v0.1.9) (2026-02-06)


### Bug Fixes

* update excel-vba.png to improve visual representation ([4261c11](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4261c11c681e04c6c2b15b367641d8f239361658))

### [0.1.8](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.7...v0.1.8) (2026-02-06)


### Bug Fixes

* enhance handling of Document Modules in Save-VBA.ps1 ([c5c35ca](https://github.com/taizod1024/vscode-excel-vba-extension/commit/c5c35ca418cb25d56ba8ec519f1e4ba1511fd69d))
* refactor code extraction from Document Module to improve metadata handling ([4f87fb6](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4f87fb656fa33601cd73de5e8052b09b4e40bbad))
* streamline Document Module handling by optimizing component retrieval and updating logic ([06e9377](https://github.com/taizod1024/vscode-excel-vba-extension/commit/06e93770626dd3824576f2f8193bde6876594fd3))
* update description in package.json to reflect CSV support ([1305b33](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1305b33c4dee5a69ef73e8fff3c7fb92855dfb71))
* update error message for Excel instance retrieval ([3e9d669](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3e9d6690d7d3816a29e18b654116f8c7360ad546))

### [0.1.7](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.6...v0.1.7) (2026-02-04)

### [0.1.6](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.5...v0.1.6) (2026-02-04)


### Bug Fixes

* enhance .url file handling and improve workbook detection in Load-CSV and Save-CSV scripts ([de7d477](https://github.com/taizod1024/vscode-excel-vba-extension/commit/de7d477a853847049a5601075c50a04e2e29a062))
* improve CSV parsing and data handling in Read-CsvFile and Update-SheetData functions ([a0387b6](https://github.com/taizod1024/vscode-excel-vba-extension/commit/a0387b65a015d7c8bfe3c2f3688b339eb2b382f6))
* rename dummy URL shortcut to URL shortcut and update related functionality ([999a4d1](https://github.com/taizod1024/vscode-excel-vba-extension/commit/999a4d1bde1653033077e3a6d054ff55b399a461))
* update error message for opened workbook check and restore Create-UrlShortcuts.ps1 script ([3ea2bad](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3ea2badc17f5f753206da66a6fd3dff59e3cf4ba))
* update placeholder text for new Excel file name input prompt ([dc731e8](https://github.com/taizod1024/vscode-excel-vba-extension/commit/dc731e8404efc89e4cc4fd016bba27ee14579735))
* update README and package.json for CSV terminology and icon changes ([dbe5b64](https://github.com/taizod1024/vscode-excel-vba-extension/commit/dbe5b6486c807d0bdcf6212c09c7ede889659f34))

### [0.1.5](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.4...v0.1.5) (2026-02-04)


### Bug Fixes

* add URL file handling and improve Excel executable path resolution in openExcelAsync method ([7aab25b](https://github.com/taizod1024/vscode-excel-vba-extension/commit/7aab25bcbeec259d8e4a99a07e4c664e162da2ef))
* remove .url file check in openExcelAsync method ([4b62b7f](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4b62b7f9d21b34f386982346245c45e7cc9ada81))

### [0.1.4](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.3...v0.1.4) (2026-02-03)


### Bug Fixes

* correct font name in Save-CSV.ps1 ([1b9db55](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1b9db5518bc163f4fe5b55f0c78c8e49f3a8192d))
* enhance macro path resolution for VBA components to locate parent Excel workbook ([3ff4de5](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3ff4de5bf727220d2c7b982e91735dfd91c210bc))
* remove unnecessary encoding comment and update variable references in scripts ([9ff76e6](https://github.com/taizod1024/vscode-excel-vba-extension/commit/9ff76e6d8f97050246df1ea14117d31dd9dcf64c))
* simplify macro path resolution logic in loadVbaAsync and compareVbaAsync methods ([5bd46be](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5bd46be21fa0d9f6bacb105a60bab81ecf1d90c0))
* update encoding comment and correct font size assignment in scripts ([fe3194c](https://github.com/taizod1024/vscode-excel-vba-extension/commit/fe3194cb2674fdc5a897e568cb6013e6be1eff6b))

### [0.1.3](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.2...v0.1.3) (2026-02-03)


### Features

* replace ([bca2f7f](https://github.com/taizod1024/vscode-excel-vba-extension/commit/bca2f7fbba2b58bb9e0ec829827dfa6a1854cd58))


### Bug Fixes

* remove unnecessary blank lines in openExcelAsync method ([1cad887](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1cad88725332147530eb737c201c9ec7afd23816))

### [0.1.2](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.1...v0.1.2) (2026-02-03)


### Features

* add command to create and open a new Excel file ([59e79fd](https://github.com/taizod1024/vscode-excel-vba-extension/commit/59e79fd3536fd8c600eb941c51cd714bc6cd57c1))
* add command to create URL shortcuts for Excel books ([ffea2b5](https://github.com/taizod1024/vscode-excel-vba-extension/commit/ffea2b50cafb8b72733cdecf2c822ec3f4612de5))
* add detailed instructions for creating a new Excel file in README.md ([bbc1259](https://github.com/taizod1024/vscode-excel-vba-extension/commit/bbc12599576cbae68210cb7fa44c4372aea7058b))
* add instructions for creating dummy URL shortcuts for cloud-hosted Excel files ([1816bac](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1816bace4b43d80037d4089642eb536966d65a2b))
* enhance URL shortcut creation for cloud-based Excel workbooks ([c729368](https://github.com/taizod1024/vscode-excel-vba-extension/commit/c729368005e2af00c86e18d11b15d237029b8723))


### Bug Fixes

* handle 2D array for single row and column exports in Load-CSV.ps1 ([157f763](https://github.com/taizod1024/vscode-excel-vba-extension/commit/157f763da6c1450c94ff23e7f84686e2c9af5706))
* improve freeze panes handling in Save-CSV.ps1 with error handling ([baed12e](https://github.com/taizod1024/vscode-excel-vba-extension/commit/baed12eb5fc8e483df2e8d397ed90dfdeffafb22))
* normalize case sensitivity in parent folder check for CSV ([20a5ff6](https://github.com/taizod1024/vscode-excel-vba-extension/commit/20a5ff6a58ffa0758111f2407e1af841f6d9b643))
* remove unnecessary return ([3dd3d38](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3dd3d38101c974f23ef092beaa9181542a140423))
* suppress output of AppActivate calls in multiple scripts ([00fd377](https://github.com/taizod1024/vscode-excel-vba-extension/commit/00fd377be6f58a69644837a9858cb4ccdb599b00))
* update table style in Save-CSV.ps1 from TableStyleLight5 to TableStyleLight2 ([23b5e9a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/23b5e9a53af192c71063ebe2f0eda211edbb1c7f))

### [0.1.1](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.1.0...v0.1.1) (2026-02-02)


### Features

* enhance CSV loading and saving functionality with directory management and macro file checks ([833cbf9](https://github.com/taizod1024/vscode-excel-vba-extension/commit/833cbf95206ef5d78996ffeb00c630612aa7d0ac))


### Bug Fixes

* update file extension conditions for command triggers in package.json ([655d2fb](https://github.com/taizod1024/vscode-excel-vba-extension/commit/655d2fb2d28a364fdeb8df20be2e3791eeaa6410))

## [0.1.0](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.17...v0.1.0) (2026-02-01)


### ⚠ BREAKING CHANGES

* ***FOLDER NAME CHANGED*** _xlsm ->_bas, _customUI -> _xml

### Features

* ***FOLDER NAME CHANGED*** _xlsm ->_bas, _customUI -> _xml ([8d7b295](https://github.com/taizod1024/vscode-excel-vba-extension/commit/8d7b29521a42f4aa278c4abda4b006fa380b6c49))

### [0.0.17](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.16...v0.0.17) (2026-02-01)

### [0.0.16](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.15...v0.0.16) (2026-02-01)


### Features

* rename macro functions for consistency and add table conversion in CSV saving ([067f8b1](https://github.com/taizod1024/vscode-excel-vba-extension/commit/067f8b16fe735dfa1a115018445c2110aceb279a))

### [0.0.15](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.14...v0.0.15) (2026-02-01)


### Features

* activate Excel window and update status bar during CSV export and import processes ([3e59b71](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3e59b71626fd81e07ca5d8d0b653cfbb7d10cecc))
* activate Excel window before disabling screen updating for improved performance ([5b087e3](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5b087e3a3cb1673a80cda33124da53e6b53bfe21))
* add CSV import and export functionality with new commands ([1b7b880](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1b7b880485464240bf132712ad0c17d29e42b1a7))
* add functionality to delete sheets that don't start with "Sheet" in Save-CSV.ps1 ([d6d8aaa](https://github.com/taizod1024/vscode-excel-vba-extension/commit/d6d8aaa19fcf92856e356d13cabb77ef02548277))
* add validation for Attribute VB_Name to ensure it matches file names in VBA files ([c04b88a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/c04b88a02b81ee5c197c36c86e6b0436100319fc))
* add validation for Attribute VB_Name to ensure it matches the file name ([93e3be7](https://github.com/taizod1024/vscode-excel-vba-extension/commit/93e3be74da1d8a82434064041d7c50d23a992730))
* disable user interaction during CSV processing for improved performance ([307980e](https://github.com/taizod1024/vscode-excel-vba-extension/commit/307980e5598d5e2f52a1fc8d0886facf0daaef54))
* enhance CSV handling and Excel file validation in Load-CSV and Save-CSV scripts ([3c05a8a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3c05a8a17ff5e49b7a985b1fc9e675cd8b6eec85))
* enhance CSV import/export by adding row count to output messages and improving sheet handling ([26e6a50](https://github.com/taizod1024/vscode-excel-vba-extension/commit/26e6a5094838a0057fb67e187f6c039fbad74dc6))
* enhance sheet processing with status updates and count tracking for improved user feedback ([4e70bae](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4e70bae2d7b736ac09a242fcc33cdfb5c75c103d))
* improve CSV export and import processes by keeping Excel open and managing alerts ([7e03ae5](https://github.com/taizod1024/vscode-excel-vba-extension/commit/7e03ae58d1206c9afdf2ae10712476e272a6d82f))
* improve CSV handling by clearing cells before populating and opening first CSV file in explorer ([4e72199](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4e721991b9e17b8810a0eeca11148bbee94e88f7))
* optimize CSV import and export processes by using array handling and disabling screen updates ([c3c37ca](https://github.com/taizod1024/vscode-excel-vba-extension/commit/c3c37caa09773400f34aa0734bc74c2aa6899616))
* refactor error handling and initialization across scripts for consistency ([30d9639](https://github.com/taizod1024/vscode-excel-vba-extension/commit/30d96396eea4caa26ffa4a2848b780b339a91445))
* refactor sheet data import process and suppress output during cell clearing ([fb6daac](https://github.com/taizod1024/vscode-excel-vba-extension/commit/fb6daac2830848eb85c5f4b614730681ac6c547d))
* rename Populate-Sheet function to Update-SheetData for clarity ([977143b](https://github.com/taizod1024/vscode-excel-vba-extension/commit/977143bd0450f0b6fb8a4afcacff6de4cd656268))
* suppress output when clearing existing sheet data for cleaner execution ([1382ec2](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1382ec23807c4ad2f7bac1d4bb50a6e8e633a567))
* update CSV data handling to use filenames directly for sheet names ([b8d6c9c](https://github.com/taizod1024/vscode-excel-vba-extension/commit/b8d6c9cab8e4aea88814e85b0e5bae646763cbe7))
* update README and package.json to clarify support for Excel Sheets and CustomUI ([7dfc9dd](https://github.com/taizod1024/vscode-excel-vba-extension/commit/7dfc9dd9b231f76e859ac0fae7ecfea24cffd09d))
* update README with new features and clarifications for CSV and VBA file handling ([08ccd0a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/08ccd0a12932686ef567bafba164970ed07135f1))
* update terminology from "Excel Macro" to "Excel Book" across documentation and scripts ([0cabd5a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/0cabd5a60c2472473769c51d84893821457a453f))


### Bug Fixes

* correct Japanese translation for saving VBA to Excel macro in README ([2875d48](https://github.com/taizod1024/vscode-excel-vba-extension/commit/2875d48fa808ac1cde6cc70c9a68f28a4359899b))
* correct typos in Japanese README for clarity ([1f63e29](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1f63e29b09d1cb98000108f785f4ce91bf4af85b))

### [0.0.14](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.13...v0.0.14) (2026-01-31)


### Features

* update README to clarify VBA file addition and CustomUI XML registration requirements ([5450f88](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5450f887deba77ca4721ee690b5f61858a65e779))

### [0.0.13](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.12...v0.0.13) (2026-01-31)


### Features

* update excel-vba.png image for improved visual representation ([bb708b6](https://github.com/taizod1024/vscode-excel-vba-extension/commit/bb708b64539dbc13a590543cac0b91cb453a7faf))

### [0.0.12](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.11...v0.0.12) (2026-01-31)


### Features

* bring Excel and VBE windows to foreground before running macros ([3ba4918](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3ba4918628732be3f2a6a56bf58d9fa1ec0ee167))
* improve handling of temporary Excel files and add warning for .xlam save limitations ([b0388de](https://github.com/taizod1024/vscode-excel-vba-extension/commit/b0388de7f0944356e5a21f751878f655c0d64e06))
* open first loaded file in explorer view after organizing files ([7afe801](https://github.com/taizod1024/vscode-excel-vba-extension/commit/7afe801c70ed9013b132655f61fd60c9adc042d1))
* update command icons and improve user feedback for running VBA subs ([c42f093](https://github.com/taizod1024/vscode-excel-vba-extension/commit/c42f093f212e12bf043acf1f45280c7d3a8d97e7))


### Bug Fixes

* remove unnecessary blank lines in resolveVbaPath method ([273ccc5](https://github.com/taizod1024/vscode-excel-vba-extension/commit/273ccc541f0893200e383749b05c732fe06ef7d0))
* update README and code to support .xlsm files for CustomUI operations ([bac7ec7](https://github.com/taizod1024/vscode-excel-vba-extension/commit/bac7ec7ab0f8445d1d679f4a4683e10e1fead883))
* update titles and conditions for CustomUI commands in package.json ([dad60b3](https://github.com/taizod1024/vscode-excel-vba-extension/commit/dad60b32d027185893e25cac5cd06a887a620058))

### [0.0.11](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.10...v0.0.11) (2026-01-29)


### Features

* add run sub at cursor functionality ([9d28d21](https://github.com/taizod1024/vscode-excel-vba-extension/commit/9d28d21175753065d61aa72bafdf2bca578b9052))
* resolve customUI XLAM file from folder name ([a6520ab](https://github.com/taizod1024/vscode-excel-vba-extension/commit/a6520ab9245ce8fa29caf3efa51f3c57807074da))
* save VBA before running sub ([338c58f](https://github.com/taizod1024/vscode-excel-vba-extension/commit/338c58fa4a41e24960e2d1ce311ca19b8564a64a))


### Bug Fixes

* clean up whitespace in ExcelVba class and README ([eb136b8](https://github.com/taizod1024/vscode-excel-vba-extension/commit/eb136b808a14f162fc5643754f63c1f1780a7edc))
* remove extra quotes from file path in openExcel ([06a27a3](https://github.com/taizod1024/vscode-excel-vba-extension/commit/06a27a388c210344145e5f5ccf013ea8f70ee71a))
* use excel.exe explicitly in openExcel command ([59a3f85](https://github.com/taizod1024/vscode-excel-vba-extension/commit/59a3f85f5f84a071f1f17f03d5464e9c6a15d465))

### [0.0.10](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.9...v0.0.10) (2026-01-28)


### Bug Fixes

* update method to activate VB Project for manual saving ([ed0ebf1](https://github.com/taizod1024/vscode-excel-vba-extension/commit/ed0ebf1278f2784f883370b87fa8934500779006))

### [0.0.9](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.8...v0.0.9) (2026-01-28)

### [0.0.8](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.7...v0.0.8) (2026-01-28)


### Features

* add VBA project compilation step with error handling ([18bc1b6](https://github.com/taizod1024/vscode-excel-vba-extension/commit/18bc1b6b812beae6e809d88e6532cd2d64a46800))


### Bug Fixes

* improve visibility of VB Editor during manual save for add-ins ([4c29804](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4c29804f65c495e8a4e6a8f827279e5441e1179a))

### [0.0.7](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.6...v0.0.7) (2026-01-28)


### Bug Fixes

* update README and Save-VBA.ps1 to clarify manual saving for .xlam add-ins ([a46de50](https://github.com/taizod1024/vscode-excel-vba-extension/commit/a46de5044249aeb45c6693875bcb6505f9364ccc))

### [0.0.6](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.5...v0.0.6) (2026-01-27)


### Bug Fixes

* update display name and description in package.json for clarity ([679f1ec](https://github.com/taizod1024/vscode-excel-vba-extension/commit/679f1ec36ebfb7c769ebf716afcd68685c71696f))

### [0.0.5](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.4...v0.0.5) (2026-01-27)


### Bug Fixes

* refine customUI XML file detection to exclude directories ([70501bf](https://github.com/taizod1024/vscode-excel-vba-extension/commit/70501bfda36a4007aad4e992a7565f8f101af623))

### [0.0.4](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.3...v0.0.4) (2026-01-27)


### Features

* enhance workbook/add-in detection and saving logic in Load-VBA and Save-VBA scripts ([b1c59e8](https://github.com/taizod1024/vscode-excel-vba-extension/commit/b1c59e87e1338ad028489ec6ca029e10b38a5e74))
* improve detection logic for workbooks and add-ins in Load-VBA and Save-VBA scripts ([80897fc](https://github.com/taizod1024/vscode-excel-vba-extension/commit/80897fc49421bcb8cdad607864ebc5be89c185f6))


### Bug Fixes

* handle errors when closing diff editors to prevent crashes ([5e17ab3](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5e17ab3c1d5c53428bf053fc658af619dbe257c6))

### [0.0.3](https://github.com/taizod1024/vscode-excel-vba-extension/compare/v0.0.2...v0.0.3) (2026-01-27)


### Features

* update display name in package.json to reflect extension functionality ([bc8dccc](https://github.com/taizod1024/vscode-excel-vba-extension/commit/bc8dccc6c613621be6a724621cc19131dd428a01))

### 0.0.2 (2026-01-27)


### Features

* add bookPath parameter to PowerShell scripts for export/import functionality ([4f136ad](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4f136ada9d807453614fcb872a756c4744ee45e8))
* add command to open Excel book and handle related functionality ([6ae08a9](https://github.com/taizod1024/vscode-excel-vba-extension/commit/6ae08a98e9efa0a02a6335b21d162a10524ea52e))
* add compare functionality for VBA modules and organize exported files ([f44ca6a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/f44ca6a06e2a26fa393956e3f177e63738f36570))
* add CustomUI load/save functionality for Excel Add-ins ([0d5c4e2](https://github.com/taizod1024/vscode-excel-vba-extension/commit/0d5c4e2f2302012e08f6e3fdd3341733e6d9c02f))
* add function to remove blank lines before VBA code ([d1b4af2](https://github.com/taizod1024/vscode-excel-vba-extension/commit/d1b4af24307eda9f2be1ce944a3bfdf69a978ed5))
* add husky commit-msg hook for commitlint integration ([422db94](https://github.com/taizod1024/vscode-excel-vba-extension/commit/422db9448257ddb2b829520cca8beca8ac17cfaf))
* add section on unblocking access for downloaded files in README ([4fe0700](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4fe0700c0bb5914c4c9764588ba1aac305e1e932))
* enhance error handling and output logging in PowerShell execution ([f2d3404](https://github.com/taizod1024/vscode-excel-vba-extension/commit/f2d3404223cbdbb052118001bf0383947d9c41c7))
* enhance export and import functionality with progress notifications ([8a2280a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/8a2280a94ebbe49012c426b84428687adfadae31))
* enhance export/import functionality with command setup and improved error handling ([7f240fc](https://github.com/taizod1024/vscode-excel-vba-extension/commit/7f240fcbdfe1d9bfd003639a5acc8f87f81eca6c))
* enhance logging and error handling in Export and Import VBA scripts ([b4b0e29](https://github.com/taizod1024/vscode-excel-vba-extension/commit/b4b0e2987d94b3b015744e8021d44c6407b5f733))
* implement Excel VBA export functionality with error handling and component export ([ee425df](https://github.com/taizod1024/vscode-excel-vba-extension/commit/ee425dfa6f623c9a33d006c38786289fa354010b))
* implement PowerShell scripts for exporting and importing VBA with temporary path handling ([ab98399](https://github.com/taizod1024/vscode-excel-vba-extension/commit/ab98399d435543ee9a1a5e0a34a6276d403d9605))
* improve error handling and logging in PowerShell scripts and TypeScript integration ([5534ed1](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5534ed108e6431547659c7a1735c4c5b4a84538c))
* improve error messages and add workbook saving functionality in Load and Save scripts ([01b570d](https://github.com/taizod1024/vscode-excel-vba-extension/commit/01b570d7844157b77ea3ad4b5169cbb901aab0e5))
* improve Export-VBA script with enhanced logging and error handling ([d166dfa](https://github.com/taizod1024/vscode-excel-vba-extension/commit/d166dfa4247d13f339b26dee5cec52ba3dd79cd5))
* initial commit of Excel VBA module import/export extension for VSCode ([3d6753c](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3d6753c7f324036331107fb1b328f3a391624d5a))
* optimize component removal and import logic in VBA script ([8b89311](https://github.com/taizod1024/vscode-excel-vba-extension/commit/8b893118e1d212dcd21655f8f68c33aa8bde1b0f))
* organize exported files by moving them to a folder named after the workbook ([81a10e9](https://github.com/taizod1024/vscode-excel-vba-extension/commit/81a10e94b8a632ef8d34b8845193233b04978dfb))
* refactor import/export functionality to save/load and update related commands and scripts ([5242dda](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5242ddafc4652f89e888a9604b69af8d261e4817))
* remove .frx files from export and show first diff ([1609dc0](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1609dc0b3607903f097bbc836d82bb2d55187586))
* remove trailing whitespace from files during load and save operations ([e660013](https://github.com/taizod1024/vscode-excel-vba-extension/commit/e660013c9d9f0d3af005ed0ed469f730436d6197))
* remove unnecessary .frx file deletion for Form components in Load-VBA script ([4eecbdc](https://github.com/taizod1024/vscode-excel-vba-extension/commit/4eecbdc8546d32926e6afab15876e5b1c45c3269))
* remove unnecessary husky script lines for cleaner commit message handling ([3d3f087](https://github.com/taizod1024/vscode-excel-vba-extension/commit/3d3f087bae195d0d740da7eb256901759975c445))
* reorganize temporary directory handling and enhance logging in export/import processes ([8e30f31](https://github.com/taizod1024/vscode-excel-vba-extension/commit/8e30f316b12771445debcdfb75333049d14e4885))
* streamline workbook open check and export process in Export-VBA script ([a37eee2](https://github.com/taizod1024/vscode-excel-vba-extension/commit/a37eee2335c5690e87b545080e9133b6aac5158d))
* update command titles and icons for loading and saving CustomUI in package.json ([a82e867](https://github.com/taizod1024/vscode-excel-vba-extension/commit/a82e8677b60d2e5aed53b12f6736e89347d26e01))
* update command titles for clarity in README and source files ([5592e3a](https://github.com/taizod1024/vscode-excel-vba-extension/commit/5592e3a910eba59dfaf65de270927529bd601eb5))
* update import/export functionality and improve error handling in scripts ([2c6f26e](https://github.com/taizod1024/vscode-excel-vba-extension/commit/2c6f26e13920b909bd24bfb5fd346acb9f1da81a))
* update README with new features and improve file encoding handling for VBA modules ([1d62317](https://github.com/taizod1024/vscode-excel-vba-extension/commit/1d62317476a48cbab35a553b03121f18d9290944))
* update terminology from "Excel Book" to "Excel Macro" in README, scripts, and commands ([c4b6584](https://github.com/taizod1024/vscode-excel-vba-extension/commit/c4b658412f4dcde5038539e506d4972c6de3284d))


### Bug Fixes

* correct comments and improve clarity in ExcelVba class ([9259eaa](https://github.com/taizod1024/vscode-excel-vba-extension/commit/9259eaac8b4c65321bc54e1a30a3f1016c504cd8))
* correct spelling of "saving" in output messages for clarity ([a925481](https://github.com/taizod1024/vscode-excel-vba-extension/commit/a925481824981a294c5aef337fa5c7c5ae3453fb))
* update excel-vba.png image to improve visual representation ([6dcb9ba](https://github.com/taizod1024/vscode-excel-vba-extension/commit/6dcb9babbc0fd1badabcf9411570091acf0d5301))
