# CaseDesk 仕様書

最終更新: 2026-04-01

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

## 5. 台帳インポートと手動マッピング

### 5.1 目的

CaseDesk は自動検出だけを前提にしない。対象ファイル側のヘッダー追加や補助列追加を要求できない台帳でも導入できるよう、CaseDesk 側で読み取り対象と表示方法を定義できる UI を持つ。

要点は以下。

- 自動検出は便利機能として残す
- 他人の台帳では手動マッピングを正式機能として扱う
- 対象ファイルの列構造は原則変更しない
- 「どう読むか」をユーザーが明示的に確定できるようにする

### 5.2 対象選択フロー

アドイン起動時のワークブックを対象に、まずシート選択から始める。

1. 対象シートをプルダウンで選ぶ
2. 選択シート内の候補テーブルまたは使用範囲を列挙する
3. 読み取り対象を確定する
4. フィールド一覧を表示する
5. 各フィールドの表示ルールを設定する
6. 設定保存後、その設定に従って一覧・詳細を構成する

自動認識に成功した場合でも、確定前に候補内容を UI で見せる。認識不能時は黙って失敗せず、手動設定画面へ落とす。

### 5.3 設定UI

対象を確定した後、フィールド一覧を表形式で表示する。各行は台帳上の1列に対応する。

各フィールドに対して少なくとも以下を設定できる。

- 表示名
- 表示 / 非表示
- 編集可 / 読取専用
- データ形式
- CaseDesk 内の役割

`データ形式` は少なくとも以下を持つ。

- 文字列
- 複数行文字列
- 数値
- 日付
- 真偽値
- 選択肢
- パス / URL

`CaseDesk 内の役割` は最低限の必須項目だけを持つ。少なくとも以下を対象にする。

- 案件ID
- 件名 / タイトル
- 状態
- ファイルキー
- 更新日時

必須役割が不足している場合は、保存不可または明示エラーにする。

### 5.4 自動検出との関係

初回は既存の自動検出を試みてよい。ただし採用条件は以下。

- 対象シート / 対象テーブルが一意に絞れる
- 必須役割の候補が十分に埋まる
- ユーザーが確認画面で確定する

上記を満たさない場合は、自動検出結果を初期値として使うだけに留め、最終的には手動マッピングで確定する。

### 5.5 保存単位

設定は「アドインを起動したブック」単位で保持する。少なくとも以下を保存対象に含める。

- 対象シート名
- 対象テーブル名または範囲識別子
- 列ごとの表示設定
- 列ごとの編集可否
- 列ごとのデータ形式
- 列ごとの役割マッピング
- 表示順

対象ファイルの列名変更・列追加・列欠落が起きた場合は、次回起動時に差分を検出し、再確認を要求する。

### 5.6 設計上の立場

この機能は例外対応ではなく、他人の台帳に礼儀正しく入るための正式な入口とする。

- 自分用の高速導線は自動検出
- 他人向けの安定導線は手動マッピング
- どちらも同じ UI フロー上で扱う

つまり、インポートは「勝手に読む機能」ではなく、「どう読むかを合意して保存する手続き」として扱う。

## 6. 通信

### 6.1 BE → FE

BE は FE ワークブックの hidden sheet に `.Value` で直接書き込む。FE 側は `Workbook_SheetChange` を契機に `CaseDeskData.LoadFromLocalSheets` を呼び、ローカルキャッシュを更新する。

### 6.2 FE → BE

現行の FE → BE 通信は `_casedesk_request` シート経由ではない。起動時や設定変更時に `CaseDeskMain.StartWorker` / `CaseDeskWorker.UpdateConfig` を呼んで構成を渡す。

`CaseDeskWorker` 内の request dispatcher は空で、稼働中のリクエスト応答仕様はない。

## 7. スキャンループ

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

## 8. hidden sheets

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

## 9. メールデータ

CaseDesk は mail 側で `manifest.csv` を前提にする。コード上でも `RefreshMailData` は `folderPath\manifest.csv` を読む。

mail manifest の列:

```text
entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text
```

`manifest.tsv` 前提の説明は旧仕様。

## 10. 案件ファイルデータ

CaseDesk は案件ファイルをオンデマンド取得しない。`WorkerInitialScan` と継続スキャンの中で `RefreshCaseFiles` を回し、`WriteCaseFilesToFE` で `_casedesk_files` に全件反映する。

folder manifest の列:

```text
item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
```

watchbox の folder profile が `source_folder` ありならコピー同期、空なら manifest-only スキャンになる。

## 11. 設定管理

- 起動時に `_casedesk_config` `_casedesk_sources` `_casedesk_fields` を Dictionary へロード
- 実行中は Dictionary を参照
- 終了時に `BeforeWorkbookClose` → `CaseDeskLib.SaveToSheets`

## 12. ログとエラー

- Change Log は `_casedesk_log` 上の `CaseDeskLog`
- 最大 5000 行
- `ErrorHandler` は処理途中ログを蓄積し、失敗時にまとめて出す

## 13. 制約

| 項目 | 決定 |
|---|---|
| WinAPI 禁止 | VBA 標準 + COM 標準のみ |
| BE 分離 | 別プロセス Excel.Application |
| 外部データ連携 | watchbox 生成 `manifest.csv` を読む |
| シート反映 | hidden sheet 直書き + SheetChange |
| ControlSource 不使用 | コードで読み書き |
| `frmCaseDesk.Visible` 直参照禁止 | `g_formLoaded` フラグで管理 |
