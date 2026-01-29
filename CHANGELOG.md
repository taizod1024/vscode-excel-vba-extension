# Changelog

All notable changes to this project will be documented in this file. See [standard-version](https://github.com/conventional-changelog/standard-version) for commit guidelines.

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
