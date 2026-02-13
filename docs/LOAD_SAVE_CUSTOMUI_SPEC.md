# CustomUI 機能仕様書 (Load / Save)

## 概要

Excel のリボンをカスタマイズする CustomUI XML ファイルを VS Code で編集する機能です。

## 機能一覧

1. **Load CustomUI from Excel Book** - Excel アドインから CustomUI XML をエクスポート
2. **Save CustomUI to Excel Book** - VS Code の CustomUI XML を Excel アドインに保存

## CustomUI とは

**リボンカスタマイズ**

Excel のリボン（メニューバー）をカスタマイズする機能。タブ、グループ、ボタンなどの部品をXML で定義します。

**対応バージョン**
- Excel 2007 以降

**XML ファイル**
- `customUI.xml` - Office 2007-2009
- `customUI14.xml` - Office 2010 以降（推奨）

## 1. Load CustomUI from Excel Book

### 概要

Excel アドイン（.xlam）に含まれる CustomUI XML をエクスポートします。

### 入力仕様

**対象ファイル**
- Excel アドイン: .xlam
- Excel マクロブック: .xlsm（CustomUI 含む場合）

### 出力仕様

**出力フォルダ**
```
{アドイン名}_拡張子/xml/
  ├── customUI.xml
  └── customUI14.xml（ある場合）
```

例: `addin.xlam` → `addin_xlam/xml/`

**ファイル形式**
- 形式: XML
- 文字エンコーディング: UTF-8
- 改行コード: LF（Unix 形式）
- インデント: 2 スペース

### customUI.xml の構造

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/customui/2006/01">
  <ribbon startFromScratch="false">
    <tabs>
      <tab id="tabCustom" label="カスタムタブ">
        <group id="grpCustom" label="カスタムグループ">
          <button id="btnHello" label="Hello" onAction="OnClickHello"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### customUI14.xml の構造

Office 2010 以降の拡張定義：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/customui/2009/07">
  <ribbon startFromScratch="false">
    <tabs>
      <tab idMso="TabHome" insertBeforeQ="TabHelp">
        <group id="grpCustom" label="カスタムグループ" insertBeforeQ="grpEditing">
          <button id="btnClick" label="クリック" imageMso="HappyFace" size="large" onAction="OnClickHandler"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### 処理フロー

1. Excel でアドインを開く（必須）
2. VS Code でアドインファイル（.xlam）を選択
3. 「Load CustomUI from Excel Book」を実行
4. 抽出処理
   - アドインの ZIP 構造を解析
   - customUI パーツを抽出
   - XML をファイル保存

### ZIP構造（内部）

Excel アドインは ZIP ファイルとして構成：

```
MyAddin.xlam
  ├── [Rels]/
  ├── customUI/
  │   ├── customUI.xml
  │   ├── customUI14.xml
  │   └── [Rels]/
  ├── word/ または xl/
  ├── _rels/
  └── [Content_Types].xml
```

### 実装詳細

**メイン処理**: `src/commands/loadCustomUI.ts`

```typescript
export async function loadCustomUIAsync(bookPath: string, context: CommandContext)
```

**PowerShell**: `bin/Load-CustomUI.ps1`

処理内容：
1. ZIP ファイルとしてアドインを開く
2. customUI.xml を検索・抽出
3. customUI14.xml を検索・抽出
4. ファイル保存

## 2. Save CustomUI to Excel Book

### 概要

VS Code で編集した CustomUI XML をアドインに保存します。

### 入力仕様

**入力ファイル**
```
{アドイン名}_拡張子/xml/
  ├── customUI.xml
  └── customUI14.xml（ある場合）
```

例: `addin.xlam` → `addin_xlam/xml/`

### 処理フロー

1. Excel でアドインを**閉じる**（重要）
2. VS Code でアドインファイルを選択
3. 「Save CustomUI to Excel Book」を実行
4. 保存処理
   - アドインの ZIP 構造を開く
   - XML ファイルの新しい内容を読み込み
   - ZIP 内のファイルを更新
   - アドインを再構築・保存

### 重要: Excel を閉じる

アドイン保存時に **Excel がファイルをロック**するため、Excel を閉じておく必要があります。

### XML 検証

保存時に以下の検証が実施されます：

1. **XML 構文チェック**
   - 未終了タグ
   - 属性の削除漏れ
   - 文字エンコーディング

2. **属性名チェック**
   - Office 2007 と 2010 の属性を区別
   - Office 2010+（customUI14.xml）: idMso, insertBeforeQ など

### 実装詳細

**メイン処理**: `src/commands/saveCustomUI.ts`

```typescript
export async function saveCustomUIAsync(bookPath: string, context: CommandContext)
```

**PowerShell**: `bin/Save-CustomUI.ps1`

処理内容：
1. アドインのロック確認
2. ZIP ファイルとして開く
3. customUI パーツを更新
4. XML 検証
5. ZIP を再構築
6. アドイン保存

## XML エディタのサポート

VS Code で CustomUI XML を編集する際に以下をご利用ください：

**推奨拡張機能**
- XML Tools
- schemastore (スキーマ自動取得)

**IntelliSense**
- Office Open XML スキーマに基づく補完

## コールバックハンドラ

CustomUI XML のボタンクリック時に呼び出される VBA コード：

```vba
' customUI.xmlで定義されたボタン:
' <button id="btnHello" label="Hello" onAction="OnClickHello"/>
'
' 対応する VBA コード:
Public Sub OnClickHello(control As IRibbonControl)
    MsgBox "Hello clicked!"
End Sub
```

**IRibbonControl パラメータ**

```vba
' control.ID: ボタン ID
' control.Tag: プロパティ値
' control.Index: インデックス（複数選択時）
```

## イメージの使用

CustomUI でボタンにイメージを設定：

```xml
<!-- 組み込みイメージを使用 -->
<button id="btnImage" label="イメージボタン" imageMso="HappyFace"/>

<!-- カスタムイメージを使用 -->
<button id="btnCustom" label="カスタム" image="CustomImage1"/>
```

**カスタムイメージの追加**

1. イメージファイル（.bmp, .jpg, .png）を準備
2. アドインの ZIP 内に customUI フォルダに追加
3. customUI.xml で参照

## エラーハンドリング

| エラー条件 | メッセージ | 対応 |
|---------|---------|------|
| XML 構文エラー | "XML parse error" | XML を修正 |
| 属性不正 | "Unknown attribute" | 属性名を確認 |
| Excel がロック | "File is locked" | Excel を閉じる |
| ZIP エラー | "Invalid archive" | アドインを修復 |

## サンプルコード

### 最小限の customUI.xml

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/customui/2006/01">
  <ribbon startFromScratch="false">
    <tabs>
      <tab id="MyTab" label="My Tab">
        <group id="MyGroup" label="My Group">
          <button id="MyButton" label="Click Me" onAction="OnClickButton"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### VBA コールバック

```vba
Public Sub OnClickButton(control As IRibbonControl)
    MsgBox "Button clicked!"
End Sub
```

## パフォーマンス

| 操作 | 処理時間 |
|------|--------|
| Load CustomUI | < 1 秒 |
| Save CustomUI | 1-3 秒（ZIP 再構築） |

## 制限事項

1. Office 2007 以降のみ対応
2. Word, PowerPoint の CustomUI は未サポート
3. イメージはアドイン内に埋め込む必要あり
4. コールバック関数は VBA で定義が必須
5. リボンのアクティビティ化は Office の実装に依存
