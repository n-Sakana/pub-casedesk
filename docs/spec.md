# CaseDesk 仕様書

最終更新: 2026-03-29

## 1. 概要

CaseDesk は Excel VBA の案件管理ツール。Excel テーブルを直接編集しつつ、watchbox が生成した mail / folder の `manifest.csv` を読み込んで案件情報を一画面で扱う。

## 2. 基本原則

- 正本は Excel テーブル
- 外部収集は watchbox 側で行い、CaseDesk は `manifest.csv` を消費する
- FE / BE は別 Excel プロセスに分ける
- フィールド検出はセルデータから行う
- Change Log は ListObject で保持し、5000 行でローテーションする
- WinAPI は使わない

## 3. 構成

### 3.1 FE / BE 分離

```text
FE: casedesk.xlsm
  ├── frmCaseDesk / frmSettings
  ├── CaseDeskMain
  ├── CaseDeskData
  ├── CaseDeskLib
  ├── ErrorHandler / FieldEditor / SheetWatcher
  └── hidden sheets

BE: 別プロセス Excel.Application
  └── CaseDeskWorker
```

### 3.2 FE 側の責務

- ワークブック内の ListObject 列挙
- レコード表示と編集
- hidden sheet からの FE キャッシュ再読込
- 変更ログ表示
- 設定のロード / 保存

### 3.3 BE 側の責務

- mail manifest の読込
- case root / folder manifest の読込
- FE 用 hidden sheet 更新
- 差分計算
- バックグラウンド走査のスケジューリング

## 4. UI

`frmCaseDesk` は 3 カラム構成。

- 左: ソース選択、フィルタ、レコード一覧
- 中央: 詳細 / Mail / Files タブ
- 右: Change Log

`frmResize` は現行実装に存在しない。リサイズハンドル `m_resizeHandle` と左右スプリッター `m_splitterLeft` `m_splitterRight` は `frmCaseDesk` が直接持つ。

`frmCaseDeskV2` / `CaseDesk_ShowPanel2` も未実装。`docs/frmCaseDeskV2-design.md` は将来構想であり、現行仕様には含めない。

## 5. 通信

### 5.1 BE → FE

BE は FE ワークブックの hidden sheet に `.Value` で直接書き込む。FE 側は `Workbook_SheetChange` を契機に `CaseDeskData.LoadFromLocalSheets` を呼び、ローカルキャッシュを更新する。

### 5.2 FE → BE

現行の FE → BE 通信は `_casedesk_request` シート経由ではない。起動時や設定変更時に `CaseDeskMain.StartWorker` / `CaseDeskWorker.UpdateConfig` を呼んで構成を渡す。

`CaseDeskWorker` 内の request dispatcher は空で、稼働中のリクエスト応答仕様はない。

## 6. スキャンループ

`CaseDeskWorker.DoScanChunk` は 1 秒周期のラウンドロビン。

```text
TASK_MAIL
  manifest.csv の更新確認
  必要時のみ再読込

TASK_CASES
  case root / case manifest.csv の更新確認
  必要時のみ再読込

TASK_WRITE
  変更があれば _casedesk_mail / _casedesk_cases / _casedesk_files / _casedesk_diff を更新
  _casedesk_signal に version を書く
```

`YieldCallback` は次回実行を予約するだけで、旧仕様のリクエスト処理は持たない。

## 7. hidden sheets

| シート | 用途 |
|---|---|
| `_casedesk_config` | 設定 KV |
| `_casedesk_sources` | ソース設定 |
| `_casedesk_fields` | フィールド設定 |
| `_casedesk_log` | Change Log |
| `_casedesk_signal` | A1:時計 / B1:version / C1:timing |
| `_casedesk_mail` | メールレコード |
| `_casedesk_mail_idx` | メール索引 |
| `_casedesk_cases` | 案件名一覧 |
| `_casedesk_files` | 案件ファイル一覧 |
| `_casedesk_diff` | mail / case 差分 |

`_casedesk_request` は現行では作成も利用もしていないため、仕様から外す。

## 8. メールデータ

CaseDesk は mail 側で `manifest.csv` を前提にする。コード上でも `RefreshMailData` は `folderPath\manifest.csv` を読む。

mail manifest の列:

```text
entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text
```

`manifest.tsv` 前提の説明は旧仕様。

## 9. 案件ファイルデータ

CaseDesk は案件ファイルをオンデマンド取得しない。`WorkerInitialScan` と継続スキャンの中で `RefreshCaseFiles` を回し、`WriteCaseFilesToFE` で `_casedesk_files` に全件反映する。

folder manifest の列:

```text
item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
```

watchbox の folder profile が `source_folder` ありならコピー同期、空なら manifest-only スキャンになる。

## 10. 設定管理

- 起動時に `_casedesk_config` `_casedesk_sources` `_casedesk_fields` を Dictionary へロード
- 実行中は Dictionary を参照
- 終了時に `BeforeWorkbookClose` → `CaseDeskLib.SaveToSheets`

## 11. ログとエラー

- Change Log は `_casedesk_log` 上の `CaseDeskLog`
- 最大 5000 行
- `ErrorHandler` は処理途中ログを蓄積し、失敗時にまとめて出す

## 12. 制約

| 項目 | 決定 |
|---|---|
| WinAPI 禁止 | VBA 標準 + COM 標準のみ |
| BE 分離 | 別プロセス Excel.Application |
| 外部データ連携 | watchbox 生成 `manifest.csv` を読む |
| シート反映 | hidden sheet 直書き + SheetChange |
| ControlSource 不使用 | コードで読み書き |
| `frmCaseDesk.Visible` 直参照禁止 | `g_formLoaded` フラグで管理 |
