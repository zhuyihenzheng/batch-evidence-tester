# 用例 YAML 速查（直接写 / 改 `cases/**/*.yaml` 时读）

> 新建用例优先写测试用例 md 再 `autotest gen`（`docs/ai/testcase-format.md`），不要手写 YAML。
> 本表是 `config.py` 许可键的完整清单：**未知键 = 硬错误**（拼错立即报错，不静默忽略）。
> **同一个 mapping 里写两次同一个键也是硬错误**（YAML 本来是后者覆盖前者 → 前一个检查静默消失）。
> 一个 YAML = 一个用例；`id` 唯一、≤31 字。`gen` 生成的 YAML（第 1 行 `# generated-by:`）要改 md 再 gen，别直接改。

```yaml
id: TC001_x                 # 省略 = 文件名
name: "正常系：…"
description: "什么样算合格（一句话，进 Excel）"
tags: [正常系]              # 子文件夹名自动追加为 tag
enabled: false              # 布尔，不加引号。false = 未レビュー / 作成中（run 不执行）
mode: auto                  # auto | manual（人手动跑 batch，分 before/after 两次采证）

setup:                      # 顺序固定：clean/remove → sql → input_files → replace_files → batches → sql_after_batches
  clean_dirs: [input_dir, {alias: error_dir, exclude: ["Backup"]}]   # 只接受 paths 的论理名
  remove_dirs: [output_dir]           # 整个文件夹删除（论理名字符串）
  sql: ["DELETE FROM T WHERE …", {file: sql/x.sql}, {inline: "…"}]   # SQL Server 可用 GO 分隔；PostgreSQL 整段 = 1 个事务（不能写 GO）
  input_files:
    - {src: input/A.csv, dest_dir: input_dir, rename: "A_{date}.csv", corrupt: {mode: flip}}
      # src 相对用例资材目录 cases/<…>/<ID>/；可为文件夹或 glob
      # corrupt: {mode: truncate|empty|flip|garbage|append, ratio: 0.5, offset: 100, content: "…"}
      #   ratio: truncate 保留的比例 / offset: flip 的位置 / content: garbage・append 写入的内容
  replace_files: [{src: config/app.config, dest_dir: batch_dir, name: App.config}]  # finally 必还原；manual 不可用
  batches: [{batch: prepare, args: ["--date", "{date}"], expected_exit_code: 0}]     # 或 - prepare
  sql_after_batches: [{file: sql/after.sql}]
  db_lock: "BEGIN TRAN; UPDATE T WITH (TABLOCKX) SET …"   # batch 执行中另一连接持锁；manual 不可用
  # PostgreSQL: db_lock: "LOCK TABLE t_order IN ACCESS EXCLUSIVE MODE"（工具自动 BEGIN；不要写 COMMIT）

snapshot:                   # 实行前后各取一次（Excel 原生表格，变化单元格标黄）
  tables:
    - {name: T_ORDER, sql: "SELECT … ORDER BY …", mask: [CUSTOMER_NM]}   # table 是 name 的别名
      # sql 可用 {date} 系和 {batch_start} / {batch_start:%Y%m%d}

execute:
  batch: spring             # settings 的 batches.<名前>；省略 = 默认 batch（或所在文件夹的 settings）
  args: ["--mode", "daily", "--date", "{date}"]   # 必须是列表
  date: "20260803"          # 本用例的 {date} 基准日

collect:
  files: [{dir: output_dir, pattern: "RESULT_*.csv", preview: true, recursive: false}]
  folder_evidence: [input_dir, output_dir]   # 或 {targets, recursive, exclude_patterns, max_entries}；[] = 不拍

review:                     # 人工确认观点（Excel 1 行 1 项，人填 OK/NG）。整份以 REVIEW / 退出码 3 封顶直到 finalize
  - "确认……"

assert:                     # 机器判定（可省略）
  exit_code: 0                              # 或 {expected: 0, manual: true}
  job_status: COMPLETED                     # Spring Batch。或 {expected: FAILED, job: orderImportJob, manual: true}
  db:
    - {table: T_ORDER, expected: expected/TC001_x/db_T_ORDER.csv, key: [ORDER_ID],
       ignore_columns: [UPDATED_AT], manual: false}   # table 必须也在 snapshot.tables；无 expected 时须 manual: true
  files:
    - {name: "結果ファイル", actual: {dir: output_dir, pattern: "RESULT_*.csv"},
       expected: expected/TC001_x/RESULT.csv, encoding: cp932, ignore_line_patterns: ['^#作成日時']}
    - {name: "処理済へ移動", exists: {dir: processed_dir, pattern: "*.csv", count: 1}}
  log: {must_contain: ["処理正常終了"], must_not_contain: ["[ERROR]"], manual: false}   # 两个列表不能都空

teardown:                   # finally 中执行；失败记 NG
  sql: [{file: sql/x_teardown.sql}]
  clean_dirs: [input_dir]
```

## 判定规则（写断言前必须知道）

| 规则 | 理由 |
|---|---|
| pattern 匹配多个文件 → NG | 多出来的异常输出不能被忽略 |
| `assert.db` 比较 expected 的全部列与全部行；actual 多余的行 = NG | 想只比部分行，用 snapshot 的 WHERE |
| SQL Server：NULL `(NULL)`、datetime `2026-08-03 09:00:00[.ffffff]`、DECIMAL 按 scale、bit `1`/`0` | 快照绝不整形值 |
| PostgreSQL：服务器的文字表现原样（= psql 的显示）。NULL `(NULL)`、boolean `t`/`f`、timestamp 小数秒去掉尾零（`.12`）、timestamptz 带 `+09`、numeric 按 scale、bytea `\x0a0b` | 同上；不做类型转换 |
| 前后空白不忽略（CHAR(n) 的填充空格也是值）；只容许 CRLF / LF 的差 | 固定长・EDI 里空白有意义；只差空白时 NG 信息会注明，修法：SQL 里 RTRIM 或期待值补空白 |
| `manual: true` 一致也只给 REVIEW；不一致仍 NG | 人工不能覆盖自动 NG |
| `mode: manual` 的 `exit_code` 一律 SKIP | 无法验证的值不发合格证 |
| `job_status` 找不到完成行 = NG；未指定 job 且多个作业 = REVIEW | Spring Boot 退出码不可信 |
| 起动行（launched）按「作业名 + 参数」与完成行配对：对象作业有没完成的执行 = NG，其他作业有 = REVIEW | 已完成的执行不能掩盖中途停下的执行 |
| `exists` 不写 count = 1 件以上即 OK | 前一用例的残留会让它误判，建议总写 count 并 clean_dirs |
| 期待值文件按 UTF-8 读；`encoding` 只指 actual（batch 输出）的编码 | |

## settings 的 batch（`config/settings*.yaml` / `cases/<grp>/settings.yaml`）

```yaml
batch:                       # 默认 batch；batches.<名前> 写与它的差分
  type: exe                  # exe（默认）| java
  exe_path: "C:/app/batch/OrderBatch.exe"         # type: exe 时必须
  java: "C:/Program Files/Eclipse Adoptium/jdk-17/bin/java.exe"   # type: java；省略 → env_vars.JAVA_HOME → JAVA_HOME → PATH
  jar: "C:/app/order-batch/order-batch.jar"       # type: java 时必须 → java [jvm_args] -jar <jar> [common_args] [args]
  jvm_args: ["-Dfile.encoding=UTF-8", "-Dstdout.encoding=UTF-8", "-Dspring.profiles.active=it"]
  common_args: ["--spring.batch.job.name=orderImportJob"]   # Boot 2 系是 spring.batch.job.names
  working_dir: "C:/app/order-batch"               # 省略 = exe / jar 所在目录
  env_vars: {JAVA_HOME: "C:/Program Files/Eclipse Adoptium/jdk-17"}
  console_encoding: utf-8                         # stdout/stderr 的解码（C# 多为 cp932）
  timeout_sec: 600
  heartbeat_sec: 10
  log_settle_sec: 5
  log_settle_min_sec: 1
  log_settle_stable_sec: 1
  log_dir: spring_log_dir                         # 该 batch 的日志目录（论理名）；省略 = paths.log_dir
  folder_evidence: {targets: [spring_input_dir, spring_output_dir]}
```

## settings 的 database（DB 的种类与连接）

另一种 DB 的键 = 硬错误（例：PostgreSQL 写 `encrypt: true` 不会加密，所以直接报错，提示用 sslmode）。

```yaml
database:                    # SQL Server（默认）：pyodbc + Microsoft ODBC Driver
  type: sqlserver            # sqlserver（默认）| postgresql
  driver: "ODBC Driver 18 for SQL Server"   # dbcheck 会列出本机已装的驱动
  server: "SQLSRV01,1433"    # 端口用逗号；命名实例 SQLSRV01\SQLEXPRESS
  database: BATCH_DB
  auth: sql                  # sql | windows
  user: sa
  password_env: AUTOTEST_DB_PASSWORD   # 密码只经环境变量
  encrypt: true
  trust_server_certificate: true
  login_timeout_sec: 15
  query_timeout_sec: 60      # 0 = 无限制
  batch_start_margin_sec: 5  # 取不到 DB 时钟时，本机时钟往前回退的秒数
```

```yaml
database:                    # PostgreSQL：内置客户端（标准库），不需要装 psycopg2 / ODBC 驱动
  type: postgresql
  server: "pg01:5432"        # host / host:port / [::1]:5432；端口也可另写 port:（默认 5432）
  port: 5432
  database: batch_db         # 区分大小写
  user: batch_app            # 必须
  password_env: AUTOTEST_DB_PASSWORD   # 认证方式由服务器 pg_hba.conf 决定：trust / password / md5 / scram-sha-256
  sslmode: prefer            # disable | prefer（默认）| require | verify-ca | verify-full
  sslrootcert: certs/ca.pem  # verify-* 用；省略 = OS 证书库。AWS RDS 用 global-bundle.pem
  timezone: Asia/Tokyo       # 会话时区（{batch_start} 用）。batch 以别的时区写 timestamp 列时与之对齐
  search_path: "app, public" # 不带 schema 的表名从这里找
  login_timeout_sec: 15
  query_timeout_sec: 60      # → statement_timeout（0 = 无限制）
  batch_start_margin_sec: 5
```
