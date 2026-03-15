# ErrorLog (VBA Error Logging Module)

Excel VBA 用のエラーログ出力および簡易スタックトレース管理モジュールです。  
処理の呼び出し履歴を追跡し、エラー発生時にログファイルへ詳細情報を出力します。

---

## Features

- 呼び出し履歴（CallStack）の自動記録
- 処理の出入り履歴（StackTrace）の保持
- エラー発生時のログファイル出力
- ログファイルサイズ監視およびローテーション
- 実行ユーザー名・ブック名など環境情報の記録
- スタックオーバーフロー検知

---

## Log File

出力ファイル名

```
errorlog.txt
```

出力場所

```
ThisWorkbook.Path
```

ログ出力例

```
==================================================
Timestamp   : 2026/03/15 12:00:00
User        : USERNAME
File        : sample.xlsm
Procedure   : Module1.Main
Description : エラー内容
Info        : 任意情報
CallStack   : Module1.Main -> Module2.Sub
StackTrace  :
    00001 [+] Module1.Main
    00002 [+] Module2.Sub
```

---

## Usage

### 1. モジュール先頭に定数定義

```vb
Private Const MODULE_NAME As String = "Module1"
```

---

### 2. 親プロシージャ

```vb
Const PROCEDURE_NAME As String = "Main"

Call ErrorLog.Initialize
Call ErrorLog.TraceListPush(MODULE_NAME, PROCEDURE_NAME)

On Error GoTo ErrorHandler

' ===== 処理 =====

CleanExit:

Call ErrorLog.TraceListPop
Exit Sub

ErrorHandler:

ErrorLog.Raise Err.Number, Err.Description
```

---

### 3. 子プロシージャ

```vb
Const PROCEDURE_NAME As String = "SubProc"

Call ErrorLog.TraceListPush(MODULE_NAME, PROCEDURE_NAME)

On Error GoTo ErrorHandler

' ===== 処理 =====

CleanExit:

Call ErrorLog.TraceListPop
Exit Sub

ErrorHandler:

Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
```

---

## Implementation Rules

以下の記述は統一してください。

| NG | OK |
|----|----|
| On Error GoTo 0 | On Error GoTo ErrorHandler |
| Exit Sub / Exit Function | GoTo CleanExit |

---

## Notes

### 再帰呼び出し

本モジュールは再帰呼び出しに対応していません。  
必要な場合は以下定数を変更するか、Stackしない仕組みに工夫して組み替えてください。

```
CALL_STACK_THRESHOLD
```

---

### On Error Resume Next 使用時

```
On Error Resume Next
```

使用後は必ず

```
On Error GoTo ErrorHandler
```

に戻してください。

---

## Stack Control

| Item | Value |
|------|------|
| CallStack Max | 30 |
| StackTrace Max | 100 |
| Log Rotation | 3MB |

---

## Log Rotation

ログファイルが 3MB を超えた場合

```
errorlog_old.txt
```

へ退避し、新しいログファイルを作成します。

---

## Optional Information

任意情報をログへ追加できます。

```vb
ErrorLog.OtherInfo = "CustomerCode:1001"
```

---

## Intended Use

- Excel 業務システム
- 長期運用 VBA
- バッチ処理
- 障害解析

---

## ライセンス
このプロジェクトは **[MITライセンス](https://opensource.org/licenses/mit-license.php)** の下で公開されています。
