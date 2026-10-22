# 测试用例 md 格式（`autotest gen` 的输入规格）

`python -m autotest gen --md <file>` 把本格式的 md 变成 `cases/**/<ID>.yaml` + 投入文件 + 初期数据 SQL + 期待值。
转换是确定性的（不调用 LLM）。**本文是唯一规格**：解析器严格按本文实现，
不认识的标签 / 行一律报错（带行号），绝不静默忽略。改格式时同步改 `src/autotest/casegen.py` 和 `tests/test_casegen.py`
（测试会解析本文第 7 节的示例）。

## 1. 文件结构

```
---                      ← 可选 front matter（YAML）
仕様書: docs/specs/x.md   ← 追溯用，写进 YAML 头注释
出力先: order            ← cases/ 下的子文件夹（= 文件夹 tag，可放 settings.yaml 选 batch）。省略 = cases/ 直下
既定クリア: [input_dir, output_dir]   ← 未写「クリア」的用例的 clean_dirs。省略 = 内置 4 个
---
# 标题 / 概要 / 观点一览表 …   ← 导入部：第一个 "## " 之前自由书写。小节用 # 或 ###，不要用 ##
                              （带属性的数据块 / ```yaml 不能放在这里 → 报错。```sql 等示例可以）
## <ID> <ケース名>        ← 每个 "## " = 1 个用例
- ラベル: 値
```

- ID：英数字 `_` `-`，≤31 字（Excel sheet 名），整个 cases/ 内唯一（大小写不区分）。
- 用例节内只允许：`- ラベル: 値`、其子项目（`  - …` / `  1. …`）、代码块、空行、`---`、`<!-- -->`。
  `<!-- -->` 要单独占行：`-->` 后面同一行接着写的内容不会被读取，所以报错。
  散文 / 表格 / `###` 都报错 —— 补充说明写进 `- 備考:`。
- 列表型标签：值写在同一行，**或**写成子项目（每项一行），不能两者都写。同一标签不能写两次。

## 2. 标签一览

| ラベル | 必须 | 写法 | 生成的 YAML |
|---|---|---|---|
| 種別 | ✓ | `正常系` `異常系` `環境不備` `手動` | tags 开头；`手動` → `mode: manual` |
| 目的 | ✓ | 一句话：什么样算合格 | `description` |
| 仕様 | | `3.1 取込処理` | description 末尾追加「（仕様: …）」 |
| タグ | | `単体, 回帰` | tags 追加 |
| 前提 / 手順 / 備考 | | 文本或子项目 | 只进 YAML 头注释 |
| クリア | | `` `input_dir`, `output_dir` `` 或 `なし` | setup.clean_dirs（省略时见 front matter；会去掉「削除」的目录） |
| 削除 | | `` `output_dir` `` | setup.remove_dirs |
| 準備SQL / 後処理SQL | | `` `DELETE FROM T_X WHERE …` `` 或 `` `sql/x.sql` `` | setup.sql / teardown.sql |
| 入力 | | `` `input/A.csv` → `input_dir` ``（→ 省略 = input_dir） | setup.input_files |
| 前置バッチ | | `` `prepare` `--date {date}` `` | setup.batches |
| DB | | `` `T_ORDER` `SELECT … ORDER BY …` ``（省略 SQL = `SELECT *`） | snapshot.tables |
| バッチ | | `` `spring` `` | execute.batch |
| 引数 | | `` `--mode daily --date {date}` ``（一个反引号，按 shell 规则拆分，禁止 `\`） | execute.args |
| 基準日 | | `20260803` | execute.date |
| 回収 | | `` `output_dir/RESULT_*.csv` `` | collect.files（preview: true） |
| 撮影 | | `` `input_dir`, `error_dir` `` 或 `なし` | collect.folder_evidence |
| 終了コード | | `0` | assert.exit_code（`手動` 不可写） |
| ジョブ状態 | | `COMPLETED` 或 ``COMPLETED `orderImportJob` `` | assert.job_status（Spring Batch） |
| ログ必須 / ログ禁止 | | `` `処理正常終了`, `取込件数=2` `` | assert.log.must_contain / must_not_contain |
| ファイル | | `` `processed_dir/ORDER_*.csv` = 1 ``（件数必须） | assert.files（exists + count） |
| 確認 | | 子项目，1 项 = 1 个人工确认观点 | review（Excel 里逐行由人填 OK/NG） |

**值的写法**

- 字面量（日志文言、パターン、SQL、引数、论理名）用反引号 `` `…` ``；多个之间用 `,` `、` 或空格分隔。
- `論理名/パターン`：第一个 `/` 前是 settings `paths` 的论理名，后面是 glob（可含 `{date}`）。
- 行尾可加 YAML flow 选项，键名就是生成项的 YAML 键：
  `入力` `{rename: "DATA_{date}.csv", corrupt: flip}` / `DB` `{mask: [CUSTOMER_NM]}` /
  `回収` `{preview: false}` / `ファイル` `{name: "…", manual: true}` / `前置バッチ` `{expected_exit_code: 0}`。

## 3. 数据块（测试数据、期待值交给 LLM 写在这里）

代码块的 info string = `语言 属性=值 …`（值含空格用 `"…"`）。必须恰好有下列之一：

| 属性 | 写到 | 同时加进 YAML |
|---|---|---|
| `input=<論理名>/<ファイル名>` 〔encoding= newline=lf\|crlf rename= corrupt=〕 | `cases/<出力先>/<ID>/input/<ファイル名>` | setup.input_files |
| `seed=<表>` + `key=<列,…>` 或 `delete=all\|none` 〔identity=on fixture=before〕 | `cases/<出力先>/<ID>/seed.sql`（DELETE + INSERT） | setup.sql（排在「準備SQL」之后） |
| `expect_db=<表>` + `key=<列,…>` 〔ignore=<列,…> fixture=after〕 | `expected/<ID>/db_<表>.csv` | assert.db；无同名「DB」时自动补 `SELECT <CSV 的列> FROM <表> ORDER BY <key>` |
| `expect_file=<論理名>/<パターン>` 〔name= encoding= ignore_line="正则"…〕 | `expected/<ID>/<name>`（UTF-8） | assert.files 内容比对 + collect.files（preview） |
| `file=<路径>` 〔encoding= newline=〕 | `expected/…`→`expected/<ID>/…`，`fixtures/…`→`fixtures/<ID>/…`，其他→`cases/<出力先>/<ID>/…` | 无 |

- `encoding`：input/file 是写出文件的编码（C# batch 多为 `cp932`）；expect_file 是 **batch 输出文件**的编码（期待值本身总写 UTF-8，比较侧按 UTF-8 读）。默认 `utf-8`。
- seed / expect_db 的内容是带表头的 CSV。**NULL 写 `(NULL)`**（与快照的 NULL 表记相同）；空栏 = 空字符串（gen 会警告）。
  SQL 的写法按 settings 的 `database.type` 生成（gen 读 `--config` 的 settings；读不到时只有 seed / expect_db 报错）：
  - SQL Server：值以 `N'…'` 传入，靠隐式转换；空栏在数值列变 0、日期列变 1900-01-01。`identity=on` → `SET IDENTITY_INSERT`。
  - PostgreSQL：值以不带类型的 `'…'` 传入（含 `\` 时 `E'…'`），由列的类型解释；空栏在数值・日期列**报错**。
    列名生成为 `"列"`（区分大小写）→ **CSV 的列名照 DDL 写（通常小文字）**。表名 `[dbo].[T]` / `#T` 不可用。
    `identity=on` → `OVERRIDING SYSTEM VALUE`（GENERATED ALWAYS 的列需要）。投入后总会把 serial / identity 的序列推进到最大值之后
    （PostgreSQL 不会自动推进，否则 batch 下一次 INSERT 会主键冲突）；序列只进不退，新行的编号依赖环境 → 期待值里 `ignore=` 掉。
- expect_db 按快照的文字化写（比较时前后空白也算值，CHAR(n) 的填充空格要写进去，或在「DB」的 SQL 里 RTRIM）：
  - SQL Server：datetime `2026-08-03 09:00:00`（有小数秒则 6 位 `.123000`）、DECIMAL 按 scale（`15000.00`）、bit `1`/`0`。
  - PostgreSQL：= psql 的显示。boolean `t`/`f`、timestamp 小数秒去尾零（`.12`）、timestamptz 带 `+09`、numeric 按 scale。
  依赖执行时刻的列放进 `ignore=`。表的全部行都会被比较（多余行 = NG）——要只比部分行，用「DB」写带 WHERE 的 SQL。
- `fixture=before/after`：同时写 offline 用的 `fixtures/<ID>/before|after_<表>.csv`（`--offline` 下代替 DB；只验证配线，不验证 batch）。
- ```` ```yaml ````（无属性）：追加标签写不了的设定（`replace_files` / `db_lock` / `sql_after_batches` / `manual: true` 的断言…）。
  与标签生成的项**冲突即报错**（不覆盖、不合并）。`id/name/description/tags/enabled/mode` 不能写。
  块内同一键写两次（例：`assert:` 下两个 `log:`）也报错，并给出两处的 md 行号。
- **代码块的内容原样写出**：制表符（TSV）、行尾空白都保留，只去掉开头 fence 的缩进量。
- **seed / expect_db 的 CSV 不能在 `,` 后面加空格**（`1, 山田` 会投入 ` 山田`）→ 报错。值本身有前后空白时用
  `" "` 括起来（`1," 山田"`）；表头的列名不能有前后空白。行尾的空白照样是值（CHAR(n) 的填充）。
- 大量数据不要贴进 md（每次读 md 都耗 token）：放文件，用「入力」/「準備SQL」引用。

## 4. 生成规则

- 生成物一律 `enabled: false`（未レビュー的印）。人确认 md 和生成物后改成 `true`。AI 不要自己改。
- 至少要有 1 个检证项（終了コード / ジョブ状態 / ログ必須 / ログ禁止 / ファイル / 確認 / expect_db / expect_file），否则报错。
- 只自动补「不写就必然不一致」的对应：expect_db → 快照，expect_file → 回收。其余不猜。
- 投入 / 回收 / ファイル判定（件数 ≥1）用到的目录没在 clean_dirs 里 → 警告（前一个用例的残留文件会造成偽 OK）。
- 一处有错 → 一个文件都不写，所有问题一次列出（`L<行号> <ID>: …`）。

## 5. 再生成与所有权

- YAML 第 1 行是生成标记（内容哈希 + 来源 md），生成的数据文件哈希记在头注释 `# generated-file:`。
- `gen` 只写新用例；已生成且与 md 有差异的显示「差分あり」，加 `--overwrite` 才更新。
- **gen 的退出码**：0 = md 全部反映到生成物；1 = 有没反映的用例（差分あり / 対象外）—— 列在「[未反映]」里，
  并且**不放进「次の手順」**（对它们 validate / run 检查的是旧 YAML，不是 md）；2 = 格式错误或写入失败。
- YAML 或数据文件被手改（哈希不一致）→「対象外」，`--overwrite` 也不覆盖。要让 md 重新生效：删掉那个 YAML（及其数据文件）再 gen。
- 例外：只把 `enabled` 改成 `true` 不算手改；但 `--overwrite` 改了内容时会回到 `false`（需再レビュー）。
- md 里删掉的用例不自动删除，gen 会提示。gen 只管理 `# generated-file:` 里列出的文件；
  人放进 `cases/<出力先>/<ID>/`、`expected/<ID>/` 的其他文件 gen 不碰。
- md 改名 / 移动后再 gen：旧路径已不存在时视为同一来源并引き継ぐ；两个现存的 md 写了同一 ID 则报错。
- 写入是全有或全无：中途失败（例：Excel 开着期待值 CSV）会全部回滚，文件一个都不变，原因和路径会显示出来。

## 6. 命令

```bash
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --dry-run   # 只检查格式，不写文件
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md             # 生成（-v 打印生成的 YAML）
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --overwrite # md 改了之后更新
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --check     # md 与生成物是否一致（不写；不一致 = 退出码 1，CI 用）
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --json      # 结果 = 1 行 JSON（errors[] 带 line / case / message）
PYTHONPATH=src python -m autotest validate --include-disabled --case <ID>   # 对未启用的草稿跑 preflight
```

## 7. 完整示例

````markdown
---
仕様書: docs/specs/order_import.md
出力先: order_import
---
# 受注取込バッチ テストケース

### 観点一覧
| 仕様 | 観点 | ケース |
|---|---|---|
| 3.1 | 正常取込・処理済移動・結果出力 | OI_001 |
| 3.2 | 金額不正 → エラー退避・ロールバック | OI_002 |
| 5.1 | Spring Boot ジョブの完了状態 | SB_001 |

## OI_001 正常系：受注 2 件を取り込む
- 種別: 正常系
- 目的: 受注 CSV の 2 件が取り込まれ、処理済へ移動し、結果ファイルが出力されること
- 仕様: 3.1 取込処理 / 3.3 結果ファイル
- 基準日: 20260803
- 引数: `--mode daily --date {date}`
- DB: `T_ORDER` `SELECT ORDER_ID, ORDER_NO, AMOUNT, STATUS FROM T_ORDER ORDER BY ORDER_ID`
- 終了コード: 0
- ログ必須: `処理開始`, `取込件数=2`, `処理正常終了`
- ログ禁止: `[ERROR]`
- ファイル:
  - `processed_dir/ORDER_*.csv` = 1
  - `error_dir/*` = 0
- 確認:
  - 結果ファイルの件数・金額が投入データと一致していること

```csv seed=T_ORDER key=ORDER_ID fixture=before
ORDER_ID,ORDER_NO,AMOUNT,STATUS
1,A0001,15000,未処理
2,A0002,8000,未処理
```

```csv input=input_dir/ORDER_20260803.csv
ORDER_NO,CUSTOMER,AMOUNT
A0001,山田商事,15000
A0002,鈴木物産,8000
```

```csv expect_db=T_ORDER key=ORDER_ID fixture=after
ORDER_ID,ORDER_NO,AMOUNT,STATUS
1,A0001,15000,完了
2,A0002,8000,完了
```

```csv expect_file=output_dir/RESULT_*.csv name=RESULT_20260803.csv
ORDER_NO,STATUS,AMOUNT
A0001,完了,15000
A0002,完了,8000
```

## OI_002 異常系：金額が負の明細を含む
- 種別: 異常系
- 目的: 金額不正の明細があるとファイルがエラーフォルダへ退避され、終了コード 1 で終わること
- 仕様: 3.2 入力チェック
- 引数: `--mode daily --date 20260804`
- 終了コード: 1
- ログ必須: `明細エラー`, `処理異常終了`
- ログ禁止: `処理正常終了`
- ファイル:
  - `error_dir/ORDER_*.csv` = 1
  - `processed_dir/*` = 0
- 回収: `error_dir/*.csv`
- 確認:
  - 実行前後の DB を比較し、正常な明細も含めて更新が残っていないこと（ロールバック）

```csv input=input_dir/ORDER_20260804.csv
ORDER_NO,CUSTOMER,AMOUNT
A0003,佐藤工業,12000
A0004,田中電機,-500
```

## SB_001 正常系：Spring Boot ジョブが COMPLETED で終わる
- 種別: 正常系
- 目的: orderImportJob が COMPLETED で終了し、結果ファイルが 1 件出力されること
- 仕様: 5.1 ジョブ構成
- バッチ: `spring`
- クリア: `spring_input_dir`, `spring_output_dir`
- 引数: `targetDate={date}`
- 基準日: 20260803
- 終了コード: 0
- ジョブ状態: COMPLETED `orderImportJob`
- ファイル: `spring_output_dir/RESULT_*.csv` = 1
- 備考: 試験用プロファイルの設定ファイルをケース専用のものに差し替えて実行する

```csv input=spring_input_dir/ORDER_20260803.csv
ORDER_NO,CUSTOMER,AMOUNT
B0001,高橋商店,3000
```

```yaml file=config/application-it.yml
spring:
  batch:
    job:
      name: orderImportJob
```

```yaml
setup:
  replace_files:
    - src: config/application-it.yml
      dest_dir: spring_config_dir
```
````
