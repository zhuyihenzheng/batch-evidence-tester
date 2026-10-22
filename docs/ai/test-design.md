# 测试设计指南（从式样书导出用例、测试数据、期待值）

这个工具是「判定别人对不对的裁判」：**本来是 NG 却报 OK（偽 OK）比漏写一个用例更糟**。
AI 生成测试时最容易在下面标 ⚠ 的地方造出永远不会失败的检查。

## 1. 观点清单（逐项对照式样书）

| 区分 | 观点 |
|---|---|
| 正常系 | 代表数据全部取込；多个文件；0 件（空文件 / 没有文件） |
| 境界值 | 件数 0/1/N；位数上限 / 上限+1；金额 0/1/最大；日期（月末・闰年・跨年度・和暦） |
| 異常系 | 形式不正（列数・编码・BOM・空行）；值不正；重复键；DB 制约违反；部分失败时的回滚范围 |
| 環境不備 | 文件夹不存在；无权限；DB 连接失败 / 锁（`db_lock`）；设定文件差し替え（`replace_files`） |
| 再実行性 | 同一输入跑两次（二重取込？）；Spring Batch 同一参数再实行 |
| 手動 | 中途停止、时机、外部系统联动等自动化不了的 → `種別: 手動` + 「手順」 |

- 1 用例 1 观点。ID：`<机能略称>_<连番>`（≤31 字），名：`正常系：…` / `異常系：…`。
- 「目的」= 合格条件一句话；「仕様」= 式样书节号（进 description，供追溯）。

## 2. 期待结果：机器判定还是人工确认

| 写成标签 / 数据块（机器判定） | 写进「確認」（人判定） |
|---|---|
| 终了代码、ジョブ状態（式样书里有对照表时） | 式样书没写清的行为、措辞的「意思」 |
| 日志文言（式样书消息一览的原文） | 证迹图像的外观、画面 |
| 文件件数和去向（処理済 / エラー / 出力） | 性能、时机、并发 |
| DB 后状态（`expect_db`）、输出文件内容（`expect_file`） | 机器比较做不到的整合性 |

⚠ 偽 OK 防止：

- 日志文言只用式样书原文。编出来的「ログ禁止」永远不会命中，等于没检查。
- 正常系也写「ログ禁止」（填式样书里的错误文言，如 `処理異常終了`）。
- 件数判定、回收、投入用到的目录必须在「クリア」里（默认 4 个已包含；Spring 等别名要自己写）。
  否则前一个用例留下的文件会让「1 件出力された」误判为 OK。gen 会对此发警告。
- `expect_db` 比较整张表的全部行：要么 seed 用 `delete=all` 控制全表，要么「DB」写 WHERE 只取本用例的行。
- 依赖执行时刻的列放进 `ignore=`，不要编造时间戳去比较。
- Spring Boot：终了代码 **和** ジョブ状態 都写。只写终了代码的话，`main` 没写
  `System.exit(SpringApplication.exit(...))` 或经 `.bat` 启动时，作业 FAILED 也会是 0。

## 3. 测试数据

- 用最少的行覆盖观点（2〜3 行通常够）；每行的用意可写在「備考」。
- seed：`key=` 按键删除再插入（只动本用例的行）；整张表要确定时用 `delete=all`。
- NULL 写 `(NULL)`；空栏 = 空字符串（SQL Server：数值列变 0、日期列变 1900-01-01；PostgreSQL：数值・日期列报错）。
  日期 `2026-08-03 09:00:00`。
- 个人信息用虚构值；需要在 Excel 遮蔽的列：「DB」加 `{mask: [列]}`。
- 用「基準日」固定 `{date}`，期待文件名和内容与它一致。
- C# batch 的输入通常是 Shift_JIS + CRLF：`encoding=cp932 newline=crlf`。不确定就问，或写进「確認」。

## 4. 期待值

- **从式样书推导，不要照抄实际运行结果**（那等于把缺陷当成基准）。推导不出 → 写「確認」并在报告里列出。
- 按快照的文字化写。SQL Server：DECIMAL 写到 scale（`15000.00`）、bit `1`/`0`、datetime 按列的精度
  （datetime2 经 pyodbc 是 6 位小数 `.123000`）。PostgreSQL：= psql 的显示（boolean `t`/`f`、
  timestamp 小数秒去尾零 `.12`、timestamptz 带 `+09`）。
- **前后空白也是值**：CHAR(n) 列会带填充空格返回 —— 期待值写满宽度，或在「DB」的 SQL 里 `RTRIM(列) AS 列`。
  只差空白时 NG 信息会注明。
- `expect_file` 是逐行比较（行首尾空白也比较，固定长的填充不能省；只容许 CRLF / LF 的差）；
  含时间戳的行用 `ignore_line="正则"` 排除。

## 5. Spring Boot / Spring Batch 专项

- 作业参数是识别参数：同一参数再跑会被拒（JobInstanceAlreadyComplete，作业根本不启动）。
  每个用例用不同的 `targetDate`，或确认作业配置了 `RunIdIncrementer`；测「再実行」时这本身就是被测对象。
- 终了代码映射（COMPLETED=0 / FAILED=1 …）看式样书；没写就只靠「ジョブ状態」。
- chunk 中途失败：已提交的 chunk 是否保留、回滚范围 → 用 `expect_db` 区分。
- 作业名属性：Boot 3 `--spring.batch.job.name=`、Boot 2 `--spring.batch.job.names=`（写错时 Boot 2 会跑全部作业，
  `job_status` 未指定 job 时会给 REVIEW）。

## 6. PostgreSQL 专项（settings 的 `database.type: postgresql`）

- 列名区分大小写：CSV 表头、`key=`、`ignore=` 照 DDL 写（通常小文字）。写成大写会报「列が無い」（42703）。
- 采番列（serial / identity）的新编号依赖环境（序列只进不退）→ 期待值 `ignore=` 掉或只比别的列；
  GENERATED ALWAYS 的列要投入值时 seed 加 `identity=on`。
- `{batch_start}` 是 DB 会话时区的时刻。batch 以别的时区写 `timestamp`（无时区）列时，settings 的
  `database.timezone` 与 batch 对齐，否则会漏掉对象行。
- 「DB」里自己写的 SQL 用 PostgreSQL 语法（`LIMIT`、`::type`、`"列"`）；`db_lock` 写 `LOCK TABLE … IN … MODE`
  或 `SELECT … FOR UPDATE`（工具会先 BEGIN，不要写 COMMIT）。

## 7. 节省 token 的写法

- 一次写完一个机能的全部用例（同一个 md），不要一个用例一次对话。
- 多个用例共用的数据放进共通 SQL 文件，用「準備SQL: `sql/xxx_common.sql`」引用，不要每个用例重复贴。
- 用例名、确认观点写短句；导入部只放观点一览，不复述式样书。
- 出错时按 gen 给的行号局部修改，不要整份重新生成。
