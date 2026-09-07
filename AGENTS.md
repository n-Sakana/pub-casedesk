# AGENTS.md — pub/casedesk

## 役割

CaseDesk は、Excelの表、watchboxが生成したメール・案件ファイルの `manifest.csv`、変更履歴を一画面で扱うVBA製品です。ローカルWindowsとExcelで動きます。

詳細は [README.md](README.md) と [docs/spec.md](docs/spec.md) を参照してください。watchboxとの境界は [pub/watchbox/README.md](../watchbox/README.md) が正本です。

## 現行アーキテクチャ

- FE: ユーザーが操作する `casedesk.xlsm`
- BE: 別プロセスの非表示 `Excel.Application`
- BEの役割: manifest走査、案件・メールの収集、FE hidden sheetへの書込み
- FEの役割: UI、設定、hidden sheetから読み込んだDictionary cache、対象tableの読書き
- 通知: BEがhidden sheetへ `.Value` を書き、FEが `Workbook_SheetChange` で受ける

「データcacheはすべてBE」という旧説明は使いません。BEにもcacheがあり、FEも受信済みデータをDictionaryへ読み込んで表示・検索します。

## 読む順番

1. [README.md](README.md)
2. [docs/spec.md](docs/spec.md)
3. `src/CaseDeskMain.bas`
4. `src/CaseDeskWorker.bas`
5. `src/CaseDeskData.bas`
6. `src/frmCaseDesk.frm`、`src/frmSettings.frm`

## build / test

```bat
samplerun.bat
build-addin.bat
build-sample.bat
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-Compile.ps1
powershell -ExecutionPolicy Bypass -File scripts/Test-Worker.ps1
```

## 変更時の原則

- 製品runtimeはVBAで保つ。build・test scriptはこの制約の外です。
- 製品VBAにWin32 API、Shell、WMI、外部helper依存を持ち込まない。
- FE / BEの別Excelプロセス構成を維持する。
- FEとBEの終了・参照解放を変更するときは、実Excelプロセスの残留まで確認する。
- `manifest.csv` のheaderを変えるときは、watchboxとCaseDeskを同時に確認する。
- hidden sheet名、event経路、FE cache、BE cacheのどれを変えたかを区別する。
- `frmCaseDeskV2` や `CaseDesk_ShowPanel2` は現行実装ではない。構想メモを製品契約に混ぜない。
- build成功だけで完成としない。通常起動、主要画面、検索、選択、保存、終了保護を実Excelで確認する。
