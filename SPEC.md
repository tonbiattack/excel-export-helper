# excel-export-helper 仕様書

## 1. 概要

`excel-export-helper` は、Windows PC にインストール済みの Microsoft Excel を利用して、Excel ファイル（主に `.xlsx`）を TSV / PDF に変換する PowerShell ベースのローカルツールである。

Google Spreadsheet 等から一度 `.xlsx` としてダウンロードしたファイルを、Google Drive へ再アップロードしたり、ブラウザ上で編集モードで開いたりせずに、ローカル環境のみで別形式へ変換することを目的とする。

## 2. 背景

業務環境では、共有されている Google Spreadsheet を直接編集モードで開くと、意図せず更新してしまうリスクがある。そのため、基本的にプレビューで閲覧する運用となっている。

一方、プレビュー画面からは TSV や PDF など他形式への変換が行いにくいため、現状は以下のような手順が発生している。

1. Google Spreadsheet を `.xlsx` としてダウンロードする
2. 自分の Google Drive にアップロードする
3. 自分の Drive 上でファイルを開く
4. TSV / PDF などへ変換する
5. 必要なファイルを再度ダウンロードする

この一連の操作を短縮し、Google Workspace API や OAuth アプリに依存せず、ローカルだけで完結させる。

## 3. 目的

- `.xlsx` から TSV / PDF への変換を簡単に行えるようにする
- Google Drive への再アップロード作業を不要にする
- Google Drive API や OAuth 認証を不要にする
- 管理者権限や追加ソフトウェアの導入を極力不要にする
- 共有元のファイルを直接編集しない安全な運用を維持する

## 4. 想定環境

- OS: Windows 10 / 11
- PowerShell: Windows PowerShell 5.1 以上、または PowerShell 7
- Microsoft Excel: デスクトップ版がインストール済み
- 管理者権限: 原則不要

## 5. 技術方針

PowerShell から Excel COM (`Excel.Application`) を利用し、Excel 自身の保存・エクスポート機能を呼び出す。

外部ライブラリや Google API は利用しない。

### 採用理由

- 会社 PC に Excel が既にインストールされている
- Excel の表示・印刷設定を利用して PDF 化できる
- TSV も Excel 標準の保存機能で生成できる
- Python / Node.js / LibreOffice 等を追加インストールする必要がない
- Google Workspace 管理者ポリシーの影響を受けにくい

## 6. MVP

### 6.1 入力

1 個の Excel ファイルを指定する。

対応拡張子:

- `.xlsx`
- `.xlsm`（マクロは実行しない）
- `.xls` は将来的な対応候補

### 6.2 出力形式

MVP では以下をサポートする。

- TSV
- PDF
- TSV + PDF

### 6.3 基本 CLI

```powershell
./excel-export.ps1 -Path "C:\work\sample.xlsx" -Format tsv
```

```powershell
./excel-export.ps1 -Path "C:\work\sample.xlsx" -Format pdf
```

```powershell
./excel-export.ps1 -Path "C:\work\sample.xlsx" -Format all
```

### 6.4 TSV 出力

Excel ブック内の各ワークシートを個別の TSV として出力する。

入力:

```text
sample.xlsx
```

シート:

```text
Users
Orders
Settings
```

出力:

```text
sample/
├─ Users.tsv
├─ Orders.tsv
└─ Settings.tsv
```

#### ファイル名ルール

ワークシート名に Windows のファイル名として利用できない文字が含まれる場合は `_` に置換する。

対象:

```text
\ / : * ? " < > |
```

同名になる場合は連番を付与する。

例:

```text
Sheet.tsv
Sheet_2.tsv
```

### 6.5 PDF 出力

ブック全体を 1 つの PDF として出力する。

```text
sample.xlsx
↓
sample.pdf
```

Excel の `ExportAsFixedFormat` を利用する。

原則として以下の Excel 側の設定を尊重する。

- 印刷範囲
- ページ方向
- 余白
- 改ページ
- 拡大縮小
- 行高
- 列幅
- セル結合
- フォント

## 7. 出力先

デフォルトでは入力ファイルと同じディレクトリへ出力する。

例:

```text
C:\work\sample.xlsx
```

PDF:

```text
C:\work\sample.pdf
```

TSV:

```text
C:\work\sample\Users.tsv
C:\work\sample\Orders.tsv
```

任意の出力先指定にも対応する。

```powershell
./excel-export.ps1 `
  -Path "C:\Downloads\sample.xlsx" `
  -Format all `
  -OutputDirectory "C:\work\export"
```

## 8. Excel COM の動作方針

Excel はバックグラウンドで起動する。

```powershell
$excel.Visible = $false
$excel.DisplayAlerts = $false
```

変換処理終了後は必ず以下を行う。

1. Workbook を閉じる
2. Excel.Application を終了する
3. COM オブジェクトを解放する

異常終了時にも Excel.exe が残存しにくいよう、`try / finally` で終了処理を保証する。

## 9. 元ファイルの保護

本ツールは元の Excel ファイルを変更しない。

- 上書き保存しない
- 数式を書き換えない
- セルを更新しない
- マクロを実行しない

原則として読み取り専用で開く。

```powershell
$workbook = $excel.Workbooks.Open(
    $path,
    0,
    $true
)
```

## 10. セキュリティ

### 10.1 マクロ

`.xlsm` を開く場合もマクロを自動実行しないことを前提とする。

可能な限り Excel の AutomationSecurity を利用してマクロ実行を抑止する。

### 10.2 外部リンク

外部リンクの更新確認ダイアログを表示しない。

可能な限りリンク更新を行わず開く。

### 10.3 ネットワーク

ツール自身は外部通信を行わない。

Google API、Microsoft Graph、外部 Web API 等は利用しない。

## 11. エラー処理

以下を明確なエラーとして扱う。

### ファイルが存在しない

```text
ERROR: File not found: C:\work\sample.xlsx
```

### 非対応形式

```text
ERROR: Unsupported file type: .csv
```

### Excel がインストールされていない

```text
ERROR: Microsoft Excel is not installed or Excel COM is unavailable.
```

### 出力ファイルが使用中

```text
ERROR: Failed to write output file. The file may be open in another application.
```

### ワークブックを開けない

```text
ERROR: Failed to open workbook: sample.xlsx
```

## 12. 終了コード

スクリプトや CI、他ツールから利用しやすいよう終了コードを定義する。

| Exit Code | 意味 |
|---:|---|
| 0 | 正常終了 |
| 1 | 一般エラー |
| 2 | 入力ファイル不正 |
| 3 | Excel 起動失敗 |
| 4 | 変換失敗 |

## 13. ログ

通常時は簡潔なログのみ表示する。

```text
[INFO] Opening: sample.xlsx
[INFO] Exporting TSV: Users.tsv
[INFO] Exporting TSV: Orders.tsv
[INFO] Exporting PDF: sample.pdf
[INFO] Done.
```

`-Verbose` 指定時のみ詳細情報を表示する。

## 14. 将来的な機能候補

### ドラッグ＆ドロップ

`.ps1` または `.bat` に Excel ファイルをドラッグ＆ドロップして変換する。

### SendTo 対応

Windows の「送る」メニューから実行できるようにする。

```text
右クリック
  ↓
送る
  ↓
Excel Export Helper
```

### Downloads フォルダ監視

Downloads フォルダを監視し、新しい `.xlsx` を検出したら変換候補として表示する。

### 複数ファイル一括変換

```powershell
./excel-export.ps1 -Path "C:\Downloads\*.xlsx" -Format all
```

### CSV 対応

```powershell
-Format csv
```

### シート指定

```powershell
-Sheet "Users"
```

### PDF のシート別出力

```text
sample/
├─ Users.pdf
├─ Orders.pdf
└─ Settings.pdf
```

### 変換後に元ファイルを削除

明示的なオプション指定時のみ許可する。

```powershell
-RemoveSource
```

デフォルトでは削除しない。

## 15. 対象外

MVP では以下を対象外とする。

- Google Drive API との連携
- Google Spreadsheet URL からの直接ダウンロード
- OAuth 認証
- Microsoft Graph API
- Excel ファイルの編集
- マクロ実行
- GUI アプリケーション
- 常駐監視
- PDF から Excel への逆変換

## 16. 想定ディレクトリ構成

```text
excel-export-helper/
├─ README.md
├─ SPEC.md
├─ excel-export.ps1
└─ tests/
   └─ ...
```

MVP では単一の PowerShell スクリプトから開始し、必要になった段階でモジュール化する。

## 17. 完了条件

MVP は以下を満たした時点で完成とする。

- `.xlsx` を指定して実行できる
- 全ワークシートを TSV として出力できる
- ブック全体を PDF として出力できる
- `tsv / pdf / all` を指定できる
- 元ファイルを変更しない
- Excel がバックグラウンドで動作する
- 正常終了後に Excel.exe が残らない
- エラー時にも Excel.exe が極力残らない
- 外部ライブラリ不要
- Google API 不要
- 管理者権限不要
