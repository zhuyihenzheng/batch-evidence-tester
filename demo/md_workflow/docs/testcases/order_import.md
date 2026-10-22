---
仕様書: docs/specs/order_import_spec.md
出力先: order_import
---
# 受注取込バッチ テストケース

`docs/specs/order_import_spec.md` から作成。生成: `python -m autotest gen --md docs/testcases/order_import.md`

### 観点一覧

| 仕様 | 観点 | ケース |
|---|---|---|
| 4.2 / 5 | 正常明細の取込・結果ファイル・処理済移動 | OI_001 |
| 4.1 / 4.3 | 明細エラー → エラー退避・エラー明細だけ STATUS=エラー | OI_002 |
| 4.1 | 金額の境界値（1 は正常、0 はエラー） | OI_003 |
| 6 | 出力フォルダ不在 → 終了コード 2・投入ファイルは未処理のまま | OI_004 |
| 4.3 | 実行中の強制終了で中途半端な状態が残らない（手動） | OI_005 |
| 6（Java） | ジョブが COMPLETED / 終了コード 0 | SB_001 |
| 6（Java） | 明細エラーでジョブが FAILED / 結果ファイルなし | SB_002 |

## OI_001 正常系：受注 2 件を取り込む
- 種別: 正常系
- 目的: 受注 CSV の 2 件が取り込まれて STATUS が完了になり、処理済へ移動し、結果ファイルが出力されること
- 仕様: 4.2 / 4.3 / 5
- 基準日: 20260803
- 引数: `--mode daily --date {date}`
- 終了コード: 0
- ログ必須: `処理開始`, `取込件数=2`, `処理正常終了`
- ログ禁止: `明細エラー`
- ファイル:
  - `processed_dir/ORDER_*.csv` = 1
  - `error_dir/*` = 0
  - `input_dir/*` = 0

```csv seed=T_ORDER key=ORDER_ID fixture=before
ORDER_ID,ORDER_NO,CUSTOMER_NM,AMOUNT,STATUS
1,A0001,山田商事,15000,未処理
2,A0002,鈴木物産,8000,未処理
```

```csv input=input_dir/ORDER_20260803.csv
ORDER_NO,CUSTOMER,AMOUNT
A0001,山田商事,15000
A0002,鈴木物産,8000
```

```csv expect_db=T_ORDER key=ORDER_ID fixture=after
ORDER_ID,ORDER_NO,CUSTOMER_NM,AMOUNT,STATUS
1,A0001,山田商事,15000,完了
2,A0002,鈴木物産,8000,完了
```

```csv expect_file=output_dir/RESULT_*.csv name=RESULT_20260803.csv
ORDER_NO,STATUS,AMOUNT
A0001,完了,15000
A0002,完了,8000
```

## OI_002 異常系：金額が負の明細を含む
- 種別: 異常系
- 目的: 金額が負の明細だけが STATUS=エラーになり、ファイルはエラーフォルダへ退避され、終了コード 1 で終わること
- 仕様: 4.1 / 4.3 / 6
- 基準日: 20260804
- 引数: `--mode daily --date {date}`
- DB: `T_ORDER` `SELECT ORDER_ID, ORDER_NO, CUSTOMER_NM, AMOUNT, STATUS FROM T_ORDER WHERE ORDER_ID IN (3, 4) ORDER BY ORDER_ID` {mask: [CUSTOMER_NM]}
- 回収: `error_dir/*.csv`
- 終了コード: 1
- ログ必須: `明細エラー A0004`, `処理異常終了`
- ログ禁止: `処理正常終了`
- ファイル:
  - `error_dir/ORDER_*.csv` = 1
  - `processed_dir/*` = 0
- 確認:
  - エラーフォルダへ退避されたファイルが、投入したファイルと同じ内容のままであること

```csv seed=T_ORDER key=ORDER_ID fixture=before
ORDER_ID,ORDER_NO,CUSTOMER_NM,AMOUNT,STATUS
3,A0003,佐藤工業,12000,未処理
4,A0004,田中電機,-500,未処理
```

```csv input=input_dir/ORDER_20260804.csv
ORDER_NO,CUSTOMER,AMOUNT
A0003,佐藤工業,12000
A0004,田中電機,-500
```

```csv expect_db=T_ORDER key=ORDER_ID fixture=after
ORDER_ID,ORDER_NO,CUSTOMER_NM,AMOUNT,STATUS
3,A0003,佐藤工業,12000,完了
4,A0004,田中電機,-500,エラー
```

```csv expect_file=output_dir/RESULT_*.csv name=RESULT_20260804.csv
ORDER_NO,STATUS,AMOUNT
A0003,完了,12000
```

## OI_003 異常系：金額の境界値（1 と 0）
- 種別: 異常系
- 目的: 金額 1 の明細は取り込まれ、金額 0 の明細は明細エラーになること
- 仕様: 4.1
- タグ: 境界値
- 基準日: 20260805
- 引数: `--mode daily --date {date}`
- 終了コード: 1
- ログ必須: `明細エラー A0006`, `取込件数=1`
- ファイル: `error_dir/ORDER_*.csv` = 1

```csv input=input_dir/ORDER_20260805.csv
ORDER_NO,CUSTOMER,AMOUNT
A0005,小林運輸,1
A0006,加藤商会,0
```

```csv expect_file=output_dir/RESULT_*.csv name=RESULT_20260805.csv
ORDER_NO,STATUS,AMOUNT
A0005,完了,1
```

## OI_004 環境不備：出力フォルダが存在しない
- 種別: 環境不備
- 目的: 出力フォルダが無いと終了コード 2 で終わり、投入ファイルは未処理のまま投入フォルダに残ること
- 仕様: 6 / 7
- 削除: `output_dir`
- 引数: `--mode daily --date 20260803`
- 終了コード: 2
- ログ必須: `出力フォルダが存在しません`, `処理異常終了`
- ファイル:
  - `input_dir/ORDER_*.csv` = 1
  - `processed_dir/*` = 0
- 確認:
  - 標準エラーに原因（出力フォルダ不在とそのパス）が出力されていること

```csv input=input_dir/ORDER_20260803.csv
ORDER_NO,CUSTOMER,AMOUNT
A0001,山田商事,15000
```

## OI_005 手動：実行中の強制終了で中途半端な状態が残らない
- 種別: 手動
- 目的: 取込の途中で batch を強制終了しても、ファイルの二重移動や結果ファイルの欠損が起きないこと
- 仕様: 4.3
- 手順:
  - `python -m autotest manual --case OI_005 --phase before` で実行前を採取する
  - batch を起動し、ログに「ファイル取込開始」が出た直後にタスクマネージャーで強制終了する
  - `python -m autotest manual --case OI_005 --phase after` で実行後を採取する
- 回収: `output_dir/*.csv`
- 確認:
  - 投入ファイルが投入・処理済・エラーのどれか 1 か所だけに存在すること
  - 結果ファイルが出力されていないか、出力されている場合は内容が欠けていないこと

```csv input=input_dir/ORDER_20260806.csv
ORDER_NO,CUSTOMER,AMOUNT
A0007,吉田建設,5000
```

## SB_001 正常系：Spring Boot ジョブが COMPLETED で終わる
- 種別: 正常系
- 目的: orderImportJob が COMPLETED・終了コード 0 で終わり、結果ファイルが出力されること
- 仕様: 2 / 6（Java 版）
- バッチ: `spring`
- クリア: `spring_input_dir`, `spring_processed_dir`, `spring_error_dir`, `spring_output_dir`
- 撮影: `spring_input_dir`, `spring_processed_dir`, `spring_output_dir`
- 基準日: 20260803
- 引数: `targetDate={date}`
- 終了コード: 0
- ジョブ状態: COMPLETED `orderImportJob`
- ログ必須: `取込件数=2`
- ファイル: `spring_processed_dir/ORDER_*.csv` = 1

```csv input=spring_input_dir/ORDER_20260803.csv
ORDER_NO,CUSTOMER,AMOUNT
B0001,高橋商店,3000
B0002,伊藤産業,4500
```

```csv expect_file=spring_output_dir/RESULT_*.csv name=RESULT_20260803.csv
ORDER_NO,STATUS,AMOUNT
B0001,完了,3000
B0002,完了,4500
```

## SB_002 異常系：明細エラーでジョブが FAILED になる
- 種別: 異常系
- 目的: 金額 0 の明細があるとジョブが FAILED・終了コード 1 で終わり、結果ファイルは出力されないこと
- 仕様: 4.1 / 5 / 6（Java 版）
- バッチ: `spring`
- クリア: `spring_input_dir`, `spring_processed_dir`, `spring_error_dir`, `spring_output_dir`
- 撮影: `spring_input_dir`, `spring_error_dir`, `spring_output_dir`
- 引数: `targetDate=20260803`
- 終了コード: 1
- ジョブ状態: FAILED `orderImportJob`
- ログ必須: `明細エラー B0004`
- ファイル:
  - `spring_error_dir/ORDER_*.csv` = 1
  - `spring_output_dir/RESULT_*.csv` = 0

```csv input=spring_input_dir/ORDER_20260803.csv
ORDER_NO,CUSTOMER,AMOUNT
B0003,渡辺物流,2000
B0004,中村製作所,0
```
