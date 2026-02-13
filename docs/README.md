# Documentation Index

Excel VBA Extension の全機能に関する仕様書とドキュメントです。

## 全機能の概要

- [ALL_FEATURES_OVERVIEW.md](./ALL_FEATURES_OVERVIEW.md)
  - 13 個の全機能の概要
  - アーキテクチャ
  - ファイル構成
  - 共通仕様

## 機能別仕様書

### ファイル作成機能

- [NEW_BOOK_SPEC.md](./NEW_BOOK_SPEC.md)
  - New Excel Book - 標準的な Excel ファイルを作成
  - New Excel Book with CustomUI - CustomUI テンプレート付きで作成

### VBA 開発機能

- [LOAD_SAVE_VBA_SPEC.md](./LOAD_SAVE_VBA_SPEC.md)
  - Load VBA - Excel から VBA コードを抽出
  - Save VBA - VS Code のコードを Excel に保存
  - Compare VBA - コード差分を表示
  - Run VBA Sub at Cursor - Sub プロシージャを実行

### CSV シート管理機能

- [LOAD_SAVE_CSV_SPEC.md](./LOAD_SAVE_CSV_SPEC.md)
  - Load CSV - Excel シートを CSV にエクスポート
  - Save CSV - CSV ファイルを Excel シートに保存

### CustomUI 開発機能

- [LOAD_SAVE_CUSTOMUI_SPEC.md](./LOAD_SAVE_CUSTOMUI_SPEC.md)
  - Load CustomUI - XML をエクスポート
  - Save CustomUI - XML を保存

### PNG エクスポート機能

- [EXPORT_SHEET_AS_PNG_SPEC.md](./EXPORT_SHEET_AS_PNG_SPEC.md)
  - Export Sheet as PNG - シートを PNG 画像にエクスポート

### クラウド対応・その他機能

- [CLOUD_SUPPORT_SPEC.md](./CLOUD_SUPPORT_SPEC.md)
  - Create URL Shortcut - クラウドホストのショートカット作成
  - Open Excel Book - Excel で開く

## コマンドリファレンス

- [COMMAND_REFERENCE.md](./COMMAND_REFERENCE.md)
  - 全 13 コマンドの詳細リファレンス
  - コマンド ID、実行方法、前提条件
  - 処理フロー、出力仕様
  - クイックスタートガイド

## クイックスタート

### VBA 開発を開始したい

1. [NEW_BOOK_SPEC.md](./NEW_BOOK_SPEC.md) を一読
2. 「New Excel Book with CustomUI」でファイルを作成
3. [LOAD_SAVE_VBA_SPEC.md](./LOAD_SAVE_VBA_SPEC.md) を参照
4. Load VBA → 編集 → Save VBA のサイクルで開発

### CSV シートを管理したい

1. [LOAD_SAVE_CSV_SPEC.md](./LOAD_SAVE_CSV_SPEC.md) を参照
2. Load CSV でエクスポート
3. Git で管理
4. Save CSV で反映

### CustomUI をカスタマイズしたい

1. [LOAD_SAVE_CUSTOMUI_SPEC.md](./LOAD_SAVE_CUSTOMUI_SPEC.md) を参照
2. Load CustomUI で XML をエクスポート
3. IntelliSense を使って編集
4. Save CustomUI で保存

### クラウドホストのファイルを管理したい

1. [CLOUD_SUPPORT_SPEC.md](./CLOUD_SUPPORT_SPEC.md) を参照
2. Create URL Shortcut でショートカット作成
3. 通常通り VBA/CSV を管理

## ドキュメント構成図

```
ALL_FEATURES_OVERVIEW.md（全体像）
    ├── NEW_BOOK_SPEC.md（ファイル作成）
    ├── LOAD_SAVE_VBA_SPEC.md（VBA 開発）
    ├── LOAD_SAVE_CSV_SPEC.md（CSV 管理）
    ├── LOAD_SAVE_CUSTOMUI_SPEC.md（CustomUI）
    ├── EXPORT_SHEET_AS_PNG_SPEC.md（PNG）
    ├── CLOUD_SUPPORT_SPEC.md（クラウド）
    └── COMMAND_REFERENCE.md（コマンド一覧）
```

## ドキュメント検索

**「〇〇がしたい」の場合**

| 目的 | 参照先 |
|------|--------|
| 新しい Excel ファイルを作成したい | NEW_BOOK_SPEC.md |
| VBA コードを編集したい | LOAD_SAVE_VBA_SPEC.md |
| Excel シートを Git で管理したい | LOAD_SAVE_CSV_SPEC.md |
| リボンをカスタマイズしたい | LOAD_SAVE_CUSTOMUI_SPEC.md |
| ダッシュボードを PNG で保存したい | EXPORT_SHEET_AS_PNG_SPEC.md |
| OneDrive のファイルを管理したい | CLOUD_SUPPORT_SPEC.md |
| 특정 コマンドの使い方を知りたい | COMMAND_REFERENCE.md |

## 技術情報

### アーキテクチャ

詳細は [ALL_FEATURES_OVERVIEW.md](./ALL_FEATURES_OVERVIEW.md#アーキテクチャ) を参照

### 実装ガイド（旧）

- [EXPORT_SHEET_AS_PNG_IMPL.md](./EXPORT_SHEET_AS_PNG_IMPL.md)
  - 実装詳細、デバッグ方法、テスト方法

## FAQ と トラブルシューティング

各仕様書の「エラーハンドリング」セクションを参照してください。

一般的な問題：

| 問題 | 参照先 |
|------|--------|
| Excel が起動しない | 各仕様書のエラーハンドリング |
| ファイルが見つからない | CLOUD_SUPPORT_SPEC.md |
| XML 構文エラー | LOAD_SAVE_CUSTOMUI_SPEC.md |
| VBA 保存に失敗 | LOAD_SAVE_VBA_SPEC.md |

## 推奨リーディング順序

初めて使う場合：

1. ALL_FEATURES_OVERVIEW.md - 全体像を理解
2. COMMAND_REFERENCE.md - コマンド一覧を確認
3. 目的に応じて個別の仕様書を参照

頻繁に使う場合：

- COMMAND_REFERENCE.md をブックマーク
- 個別機能は必要な時に仕様書で確認

