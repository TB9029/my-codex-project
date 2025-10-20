# 在庫チェック（GS1スキャン）システム

Google Apps Script (GAS) と Google スプレッドシートで構築するモバイル優先の在庫チェックシステムです。GS1 バーコードのスキャンによる在庫確認、マスタ管理、商品定数の月次モニタリング、CSV 連携を 1 つの WebApp で提供します。

## 機能ハイライト

- 📷 **在庫チェック業務**: ZXing ブラウザ SDK を用いてスマホ/タブレットのカメラから GS1 を読み取り。期限判定（🚨/⚠️/✅/OK）を即時表示し、連続スキャン向けに 0.3 秒以内の再スキャン抑止や手入力モード、オフラインキュー送信を実装。
- 🏠 **ダッシュボード**: 「期限切れ」「1か月未満」「2か月未満」を集計しカード表示。最新 10 件のスキャン履歴、タグ・GS1 を横断する検索機能を備えています。
- 🗂️ **マスタ管理**: 商品登録フォーム、タグフィルタ付き一覧、モーダル編集（有効/無効切替）。CSV インポート/エクスポートに対応し、ヘッダ固定・UTF-8 (BOM なし) で入出力します。
- 📊 **商品定数**: 各商品ごとの「定数×QTY」基準と月次在庫数を可視化。基準未達を赤、基準一致を黄、基準超過を緑で表示し、不足行のみの CSV 出力が可能です。
- 🧠 **GS1 ユーティリティ**: (01) GTIN14、(17) 期限、(10) ロット、(21) シリアルを解析し、期限ルールに基づいたアイコン／色を決定。日付は Asia/Tokyo に正規化します。
- ⚙️ **GAS 最適化**: WebApp 側の HTML テンプレートでは `HtmlService.createTemplateFromFile` を利用し、GAS で禁止されている `addMetaTag` 等の API は使用していません。

## ファイル構成

```
├── Code.gs            # GAS バックエンド（doGet, API, GS1 パーサ, CRUD）
├── Index.html         # WebApp エントリ。Styles/Scripts パーシャルを読み込み
├── Styles.html        # <style> を含むモバイル優先 UI（ライト/ダーク自動）
├── Scripts.html       # <script> を含むフロントロジック（SPA, ZXing, offline queue）
├── README.md          # 本ドキュメント
├── .gitignore
├── .editorconfig
└── .github/
    └── pull_request_template.md
```

> GAS では .css/.js ファイルを直接ホストできないため、スタイルとスクリプトは HTML ファイルに内包し、`<?!= include('Styles'); ?>` のように読み込みます。

## シートスキーマ

単一のスプレッドシートに以下の 3 シートを作成します（UTF-8、日本語ヘッダ、1 行目固定）。

### `master`

| 列 | 説明 |
| --- | --- |
| id | UUID（自動採番）|
| 商品名 | 商品名 |
| GS1コード | メイン GS1 |
| 予備GS1コード1 | 予備 GS1 |
| 予備GS1コード2 | 予備 GS1 |
| 定数 | 最低必要数 |
| 単位 | 個/箱 |
| QTY | 分割数（個=1, 箱=2以上）|
| タグ | カンマ区切り |
| 作成日 | ISO 日時 |
| 更新日 | ISO 日時 |
| 有効 | TRUE/FALSE |

### `scans`

| 列 | 説明 |
| --- | --- |
| timestamp | スキャン日時 |
| raw | 読み取りテキスト |
| gtin | (01) GTIN14 |
| expiry | (17) → yyyy-MM-dd |
| lot | (10) 値 |
| serial | (21) 値 |
| マスタid | `master.id` |
| 判定 | 🚨/⚠️/✅/OK |
| 備考 | エラー概要 |
| ユーザー | `Session.getActiveUser().getEmail()` |

### `constants`

| 列 | 説明 |
| --- | --- |
| id | `master.id` |
| 商品名 | 商品名 |
| 単位 | 単位 |
| QTY | QTY |
| 定数 | 最低必要数 |
| 1月〜12月 | 月次在庫数 |
| 最終更新 | ISO 日時 |

## セットアップ手順

1. Google スプレッドシートを作成し、上記 3 シートとヘッダを用意します。
2. GAS プロジェクトを作成し、本リポジトリのファイルをコピーまたは `clasp` 等で同期します。
3. `Code.gs` を実行できる権限で開き、以下の関数を実行してスプレッドシート ID を登録します。
   ```javascript
   setSpreadsheetId('ここにスプレッドシートID');
   ```
4. WebApp をデプロイ（新しいデプロイ → 種別：Web アプリ）し、実行ユーザーを「自分」、アクセス権を「ドメイン/特定ユーザー」に設定します。
5. 初期データが必要な場合は `appendDemoScans(count)` を実行するとダミースキャンを追加できます。

## 主要 GAS 関数

- `doGet()` — WebApp の HTML を返却。
- `getAppBootstrap()` — ダッシュボード、マスタ一覧、商品定数をまとめて取得。
- `recordScan({ raw })` — GS1 を解析し、`scans` シートへログ保存。期限判定も返却。
- `createMaster(payload)` / `updateMaster(payload)` / `deleteMaster({ id })` — マスタ CRUD。
- `listMasters()` / `importMastersCsv(text)` / `exportMastersCsv()` — マスタ一覧と CSV 連携。
- `listConstants()` / `updateConstant({ id, month, value })` / `exportConstantsCsv()` / `importConstantsCsv()` / `exportDeficitCsv()` — 商品定数管理。
- `getDashboardSnapshot()` — 日切れ間近カウントと最新スキャン。

GS1 パーサ (`parseGs1_`) は (01)/(17)/(10)/(21) をサポートし、期限形式が不正な場合はエラーを返します。

## UI 操作メモ

- 画面下部の固定ナビで 4 つの主要ビューを切り替えます。マスタ管理はタブで「新規登録」「編集」を切替。
- スキャン結果は色付きカードで表示。オフライン時はローカルキューに退避され、オンライン復帰後「⏫ 送信」で一括送信します。
- マスタ一覧の行をタップするとモーダルが開き、その場で編集・削除・有効切り替えが可能です。
- 商品定数の各行には当月値を編集するボタンがあり、入力値は GAS 経由で保存されます。

## デプロイ & 運用上の注意

- GAS のログには個人情報を保存しない方針です。`scans` の備考欄には必要最小限のエラー概要のみを記録します。
- `Session.getActiveUser().getEmail()` を利用するため、組織内/特定ユーザーのみアクセス可能な設定にしてください。
- 期限判定は Asia/Tokyo タイムゾーンで行い、(17) の YYMMDD が不正な場合はエラーとして扱います。
- CSV は UTF-8 (BOM なし) を出力します。Excel で開く際はインポートウィザードを利用してください。

## ライセンス

本リポジトリのサンプルコードは MIT ライセンスで提供します。
