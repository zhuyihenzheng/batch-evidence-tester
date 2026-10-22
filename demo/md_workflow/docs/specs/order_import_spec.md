# 受注取込バッチ（ORDER_IMPORT）処理仕様書

## 1. 概要

投入フォルダの受注 CSV を読み込み、受注テーブル `T_ORDER` の状態を更新し、結果ファイルを出力する。
同じ処理を C# 版（`OrderImport.exe`）と Java 版（Spring Boot ジョブ `orderImportJob`）で提供する。

## 2. 起動方法

| 版 | 起動 |
|---|---|
| C# 版 | `OrderImport.exe --mode daily --date <業務日付 YYYYMMDD>` |
| Java 版 | `java -jar order-batch.jar --spring.batch.job.name=orderImportJob targetDate=<YYYYMMDD>` |

## 3. 入力

- 場所: 投入フォルダ（C# 版 `in/`、Java 版 `spring/in/`）
- ファイル名: `ORDER_<任意>.csv`、文字コード UTF-8、ヘッダ行あり
- 項目: `ORDER_NO`（受注番号）, `CUSTOMER`（顧客名）, `AMOUNT`（金額）

## 4. 処理

### 4.1 明細チェック
`AMOUNT` が 1 以上の整数であること。満たさない明細は「明細エラー」としてログに出力し、
`T_ORDER.STATUS` を `エラー` にする。

### 4.2 取込
チェックを通った明細は `T_ORDER.STATUS` を `完了` に更新する。

### 4.3 ファイル移動
エラー明細を 1 件でも含むファイルはエラーフォルダ（`error/`）へ、それ以外は処理済フォルダ（`processed/`）へ移動する。

## 5. 出力

- 結果ファイル: 出力フォルダ `out/RESULT_<業務日付>.csv`（UTF-8）
- 項目: `ORDER_NO`, `STATUS`（`完了`）, `AMOUNT`。チェックを通った明細のみを出力する
- Java 版は明細エラーが 1 件でもあれば結果ファイルを出力しない

## 6. 終了コード・ジョブ状態

| 条件 | C# 版 終了コード | Java 版 ジョブ状態 / 終了コード |
|---|---|---|
| 全明細正常 | 0 | COMPLETED / 0 |
| 明細エラーあり | 1 | FAILED / 1 |
| 入出力フォルダが存在しない | 2 | FAILED / 1 |

## 7. ログ

C# 版: `log/batch_<業務日付>.log`、Java 版: `spring/log/order-batch.log`

| メッセージ | 出力契機 |
|---|---|
| `処理開始` | 起動時（C# 版） |
| `明細エラー <受注番号> <理由>` | 明細チェック NG |
| `取込件数=<件数>` | 終了時 |
| `処理正常終了` / `処理異常終了` | 終了時（C# 版） |
| `<入力/出力/処理済/エラー>フォルダが存在しません: <パス>` | フォルダ不在（C# 版） |
