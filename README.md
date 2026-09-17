# Excel Export Helper

インストール済みのデスクトップ版 Microsoft Excel を利用し、Excel ブックをローカルで TSV / PDF に変換する PowerShell ツールです。Google Drive へのアップロード、Google API、OAuth、管理者権限は不要です。

## 前提条件

- Windows 10 / 11
- Windows PowerShell 5.1 以降、または PowerShell 7
- デスクトップ版 Microsoft Excel

対応する入力形式は `.xlsx` と `.xlsm` です。`.xlsm` もマクロを実行せずに開きます。

## 使い方

必要な場合だけ、その PowerShell セッションでスクリプト実行を許可します。

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

TSV のみを出力します。

```powershell
.\excel-export.ps1 -Path "C:\work\sample.xlsx" -Format tsv
```

PDF のみを出力します。

```powershell
.\excel-export.ps1 -Path "C:\work\sample.xlsx" -Format pdf
```

両方を任意のフォルダへ出力します。

```powershell
.\excel-export.ps1 -Path "C:\work\sample.xlsx" -Format all -OutputDirectory "C:\work\export"
```

## 出力先

`-OutputDirectory` を省略した場合は入力ファイルと同じフォルダです。

```text
C:\work\sample.xlsx
↓
C:\work\sample.pdf
C:\work\sample\Users.tsv
C:\work\sample\Orders.tsv
```

ワークシート名の `\ / : * ? " < > |` はファイル名では `_` に置き換えます。置き換え後に同名になる場合は `_2`、`_3` のような連番を付けます。

## 安全性

- 元ファイルは読み取り専用で開き、保存・編集しません。
- Excel は非表示で起動します。
- マクロの自動実行を抑止します。
- 外部リンクを更新せずに開きます。
- ツール自身はネットワーク通信をしません。
- TSV/PDF の成否にかかわらず、ブックと Excel COM オブジェクトを終了・解放します。

PDF はブックに設定済みの印刷範囲、余白、方向、改ページ、拡大縮小などをそのまま利用します。

## 終了コードとエラー

| Code | Meaning |
| ---: | --- |
| 0 | 正常終了 |
| 1 | 一般エラー |
| 2 | 入力ファイルまたは引数が不正 |
| 3 | Excel が未インストール、または COM を起動できない |
| 4 | ブックを開けない、または変換に失敗 |

よくあるエラーは、入力ファイルがない、`.csv` など未対応形式を指定した、出力先ファイルを別のアプリで開いている、または Excel が利用できない場合です。

## 確認手順

Microsoft Excel がある Windows PC で、次を確認してください。

1. 複数シートの `.xlsx` を `-Format all` で変換し、PDF と各シートの TSV が作成される。
2. `.xlsm` を変換し、マクロが実行されない。
3. 変換前後で入力ファイルのハッシュを比較し、一致する。
4. 出力 PDF の印刷設定が元ブックと一致する。
5. 出力 PDF/TSV を別アプリで開いた状態で変換を試し、失敗後にも余分な `EXCEL.EXE` が残らない。

ハッシュ確認の例:

```powershell
Get-FileHash "C:\work\sample.xlsx"
```

## 自動テスト

Pester が導入済みの環境では、Excel を起動せずに入力検証・パス生成・ファイル名処理を確認できます。

```powershell
Invoke-Pester -Path .\tests\ExcelExportHelper.Tests.ps1
```
