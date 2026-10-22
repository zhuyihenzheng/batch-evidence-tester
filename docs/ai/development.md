# 开发指南（改 `src/` / `tests/` / `demo/` / `tools/` 之前读）

> 从原 CLAUDE.md 迁来的工程细节（内容未删减），供所有 AI 代理与开发者按需读取。
> 常驻上下文只有 `AGENTS.md`；这里是「动代码时才需要」的部分。

## 硬性约束（细节）

### Python 3.6 语法兼容

对象环境含 Anaconda 5.2（Python 3.6.5 / PyYAML 3.12 / openpyxl 2.5），**3.7+ 的语法/API 会在 import 时直接炸掉**。
所以：不用 dataclass（`models.py` 全是手写 `__init__`）、不用 `subprocess(capture_output=/text=)`、
不用 `Path.is_relative_to`（用 `fsops._is_relative_to`）、类型标注用 `typing.List` 而非 `list[...]`、
不用 f-string `=` 说明符 / 海象运算符 / `str.removeprefix`。PyYAML 3.12 没有 `dump(sort_keys=)`
（`casegen.py` 因此自己拼 YAML 文本，并读回校验）。

改动 `src/` 或 `demo/` 后必须过这一关：

```bash
python tools/check_py36_compat.py src/autotest demo
```

版本分支（`if sys.version_info >= (3, 12):`）是被允许的，检查器认识这种守卫。
检查器只认识它列出的模式；手边有旧解释器时，可用 Python 3.8 + `pyyaml==3.13` + `openpyxl==2.5.14`
的 venv 跑一遍 `tests.test_casegen tests.test_java_batch tests.test_regressions` 做近似确认。

### 不走 `pip install -e .`

社内封闭网络下构建隔离会去联网取 setuptools 而失败。所以 CLI 一律靠 `PYTHONPATH=src` 启动，
依赖只用 `requirements.txt`。`.bat` 脚本已内置这一行；手敲命令时要自己带上。

## 常用命令

```bash
# 单元测试（全部写成标准库 unittest 风格；pytest 不在 requirements 里，本地不需要装）
python -m unittest discover -s tests -v
python -m unittest tests.test_regressions.TestKeepEnvOption.test_teardown_skipped_and_reported  # 单个
#   ※ 唯一的例外是 GitHub Actions，它另装 pytest 跑 layout 相关测试（见「EXE 打包与 CI」）。
#      所以测试要保持 unittest 风格 —— 这样两边都能跑。

# 3.6 兼容检查
python tools/check_py36_compat.py src/autotest demo

# 开发主循环：不碰真 DB / 真 exe 的沙箱全流程（macOS 上也能跑）
PYTHONPATH=src python -m autotest run --config config/settings.demo.yaml --offline
#   → demo/fake_batch.py 当 .exe、demo/fake_spring_batch.py 当 Spring Boot jar，fixtures/ 当 DB
#   → TC003 是「故意做成 NG 的样本」，所以这条命令正常情况下退出码就是 1

PYTHONPATH=src python -m autotest list                # 用例和 tag 一览（--include-disabled 含未启用）
PYTHONPATH=src python -m autotest validate            # 设定 + 全用例 preflight（有问题退 1）
PYTHONPATH=src python -m autotest dbcheck             # 实际连一次 DB（SQL Server 需 Windows/pyodbc；PostgreSQL 哪里都能跑）
#   PostgreSQL 实服务器测试（平时 skip）:
#   AUTOTEST_PG_TEST_URL=postgresql://user:pw@127.0.0.1:5432/db python -m unittest tests.test_postgresql_live
PYTHONPATH=src python -m autotest run --dry-run       # 不启动 exe、不连 DB，只验流程

# 用例作成 / 手动实施 / 证迹结合
PYTHONPATH=src python -m autotest new --id TC010_x --template normal|error|env_missing|manual
PYTHONPATH=src python -m autotest copy --from TC001_normal --id TC011_x
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md [--dry-run] [--overwrite] [--check] [--json]
PYTHONPATH=src python -m autotest manual --case TC008_manual_demo --phase before|after
PYTHONPATH=src python -m autotest report --run <run_id> --run <session> --out merged.xlsx
PYTHONPATH=src python -m autotest finalize --run <run_id> --run <session> --excel merged.xlsx --out final.xlsx
PYTHONPATH=src python -m autotest gui                 # 操作画面（需要 tkinter）

# md 工作流样例（式样书 → 测试用例 md → gen → run）
cd demo/md_workflow && PYTHONPATH=../../src python -m autotest gen --md docs/testcases/order_import.md --config config/settings.yaml

# layout_txt（另一个工具，和上面的用例体系无关）
PYTHONPATH=src python -m autotest.layout_txt definition.xlsx --out-dir output/layout_txt
PYTHONPATH=src python -m autotest.layout_txt --gui     # 独立 GUI

# 转送包：无法传文件的环境靠 RDP 剪贴板搬项目（生成自解压 .ps1 到 transfer/）
python tools/make_transfer_bundle.py --chunk-kb 30
```

GUI 的自动化冒烟测试要 `root.withdraw()` + 只用 `root.update()` 驱动事件循环，
不要 `mainloop()`（会挂住）。

Windows 侧入口是 `.bat`（`setup_windows.bat` / `run_test.bat` / `run_demo.bat` /
`run_layout_txt.bat` / `build_layout_exe.bat`），它们会自动选 `.venv` 或 PATH 上的 python
并设好 `PYTHONPATH`。macOS 上直接用上面的 CLI。

**退出码是对外契约**（タスクスケジューラ / CI 依赖它）：`0`=全 OK、`1`=有 NG、`2`=设定错误、`3`=有「要確認」。

## 架构要点

一个用例的完整生命周期集中在 `orchestrator.CaseRunner.run()` 一个方法里，顺序本身是设计（模块头部有注释说明为什么）：
preflight → setup（清空/SQL/投入）→ 実行前快照+文件夹撮影 → 记录日志字节偏移 → 跑 batch →
等日志写入落定 → 回收成果 → 実行後快照+撮影 → 渲染证据图 → 判定 → finally 里 teardown/还原。

- **`models.py` 是唯一的层间契约**。`excel.py` 只认识 `Table` / `CheckResult` / `CaseResult` / `RunResult`，
  不知道数据来自 DB、文件还是日志。加新证据类型时优先复用这几个类型，别让 excel.py 去认识新的来源。
- **文件夹论理名（path alias）是环境抽象**。用例只写 `input_dir`，物理路径只存在于 `settings.yaml` 的 `paths:`。
  `clean_dirs` / `remove_dirs` / `collect.files.dir` 只接受论理名——这是防误删的第一道闸。
- **配置文件自动选取**：存在 `config/settings.local.yaml` 就优先用它（已 gitignore，放各人的实路径和 DB 连接），
  否则用 `settings.yaml`（仓库维护的模板）。所以不要在 `run_test.bat` 里写死 `--config`。
  项目根 = 配置文件所在 `config/` 的上一级（`cli._project_root_for`；`cases/` 还不存在时也成立——
  否则 `gen` 会把新项目的生成物写进工具自己的仓库）。
- **batch 的起动方式**（`runner.build_command` 一处组装，运行 / preflight / 证迹显示共用）：
  `type: exe`（默认）= `exe_path + common_args + execute.args`；
  `type: java` = `java [jvm_args] -jar <jar> [common_args] [execute.args]`。java 的解析顺序：
  `batch.java` → `batch.env_vars.JAVA_HOME` → 环境变量 `JAVA_HOME` → PATH（validate 会显示实际解析结果）。
  jar 按项目根解析成绝对路径再传（working_dir 不同也不会指错）。`BATCH_PROFILE_KEYS` 之外的键是硬错误。
- **Spring Boot 的判定**：`assert.job_status` 从 stdout/stderr/日志切片里读 Spring Batch 启动器的完成行
  `Job: [...] completed ... status: [X]`。没有完成行 = NG（成功的证据不存在）；未指定 `job` 而检测到多个
  作业 = REVIEW（Boot 2 忘写 `spring.batch.job.names` 会跑全部作业）。`validate` 对「java batch 只断言退出码」给出推荐警告。
- **DB 的种类**（`database.type`，`config.DATABASE_*_KEYS` 按种类限定可写的键）：
  `sqlserver`（默认）= `PyodbcClient`；`postgresql` = `PostgresClient` + `pgwire.py`。
  pgwire 是标准库写的协议 v3 客户端（闭域网里装不了 psycopg2 / psqlODBC，文本能随项目带进去）：
  只用 Simple Query；**值不做类型转换**（NULL → None，其余 = 服务器的文字表现，与 psql 一致）；
  起动时固定 `client_encoding / DateStyle / bytea_output / extra_float_digits`；认证 trust / password /
  md5 / SCRAM-SHA-256（含 SASLprep），其余报错；`sslmode` 5 种，`require` 被拒绝时不降级为明文。
  连接一律经 `db.connect_postgres`（查询 / `LockHolder` / `dbcheck` 共用）。方言差异集中在：
  `NOW_QUERIES`（`SYSDATETIME()` vs `clock_timestamp()`——`now()` 停在事务开始时刻）、
  `format_sql_datetime`（毫秒 vs 微秒）、`LockHolder`（PG 用 ReadyForQuery 的事务状态确认锁还在）、
  `diagnose_pg_error`（按 SQLSTATE，不靠随 lc_messages 变化的文言）、gen 的 `SqlServerDialect` / `PostgresDialect`。
- **offline 模式**：`fixtures/<case_id>/<before|after>_<TABLE>.csv` 顶替 DB 快照，setup SQL 不执行
  （所以靠 SQL 注入故障的用例在 offline 下无效，它们 `enabled: false`）。
- **期待值基线流程**：可以从式样书推导期待值（`gen` 的数据块），也可以跑通一次 → 从 `output/<run_id>/artifacts/`
  和 Excel 的「実行後」表取实际值 → 人工确认 → 放进 `expected/<case_id>/` 当基线。
- **`mode: manual` 把生命周期劈成两个进程**：`run_manual_before` / `run_manual_after`
  复用 `run()` 的同一批私有方法（生命周期知识只有一份）。跨进程状态存
  `output/manual_<id>_<ts>/session.json`，字段是固定 schema，`format_version` 不匹配直接报错。
  刻意**不**持久化 `replace_files` 退避信息和 `db_lock`——它们依赖 `run()` 的 `finally` 还原，
  两个进程之间没有 finally，所以 preflight 直接禁止在 manual 用例里使用。
  `base_date` 必须从 session 恢复（用当天日期会让跨日的 `{date}` 静默错配）。
- **GUI 是 CLI 之上的一层**：`gui.py` 一律 `subprocess` 调 `python -m autotest ...`，
  绝不在进程内调 `CaseRunner`。原因：长跑/pyodbc 崩溃不能拖死界面；tkinter 非线程安全；
  编排入口保持一条（测试守的是 CLI 这条路）。进度靠守护线程读 stdout → `queue.Queue`
  → `root.after()` 排水，Tk 组件只在主线程碰。
- **用例发现是递归的**：`cases/group_a/subgroup_1/TC001.yaml` 自动带上
  tag `[group_a, subgroup_1]`；
  `cases/TC001/` 这种「与定义文件同名的目录」被视为该用例的资材目录，其中的 YAML 不会被当成用例。
- **一个 YAML = 一个用例**，没有第二种格式。曾经做过「一个文件多用例（`defaults:` + `cases:`）」，
  已删除——它和 `autotest copy` / GUI 的複製按钮互斥，而且深合并的「列表整个替换」会让
  case 级的 `assert.db` 静默顶掉 defaults 继承的检查（断言变少、零警告，正是偽 OK 方向）。
  要共享共通部分，走**显式的按文件引用**（如 `sql: {file:}`），不要做隐式合并。
  测试用例 md（`gen` 的输入）不是第二种用例格式：它被确定性地展开成一个个 YAML，运行时只认 YAML。

## gen（测试用例 md → YAML）子系统

`casegen.py`。规格是 `docs/ai/testcase-format.md`（**唯一规格**，测试会解析其第 7 节示例并检查标签/属性全部出现在文档里）。

- 未知标签、无法解释的行、标签与 ```yaml 块的同一键冲突 → 全部报错（带行号），**一处错就一个文件都不写**。
  YAML 块与标签冲突时不覆盖也不合并（理由同上：静默丢断言）。
- 生成物一律 `enabled: false`；没有任何检证项的用例拒绝生成；`手動` 写 `終了コード` 拒绝（会被 SKIP 却像确认过）。
- 所有权：YAML 第 1 行 `# generated-by: autotest gen sha256=<正文哈希> md=<来源>`，生成的数据文件哈希记在
  `# generated-file:` 行。哈希不一致（手改过）= 永不覆盖。哈希对 CRLF 归一（Windows git 的 autocrlf），
  并忽略 `enabled:` 行（レビュー后改 true 不算手改；`--overwrite` 改了内容会重置为 false）。
  因此**头注释里不要出现 `enabled: false` 字面量**（编辑器一括置換会改到注释导致误判）。
- seed 数据 → `seed.sql`（`DELETE`（按 key / 全部 / 不删，必须显式）+ 1000 行一批的 `INSERT`，
  `(NULL)` → NULL，与 `db.NULL_TEXT` 一致）。expect_db 无同名「DB」时补 `SELECT <期待 CSV 的列> … ORDER BY <key>`。
  写法由 `database.type` 决定（`casegen.DIALECTS`；gen 读 `--config` 的 settings，读不到时只有 SQL 块报错）：
  SQL Server = `N'…'` / `[列]` / `IDENTITY_INSERT` / `GO`；PostgreSQL = `'…'`（含 `\` 时 `E'…'`）/ `"列"` /
  `OVERRIDING SYSTEM VALUE` / 投入后推进 serial・identity 的序列（DO 块，只进不退）。
  **PostgreSQL 不能用 `N'…'`**：会变成 character 型，投入 varchar / text 时末尾空白被吞（实测）。
  SQL Server 的输出必须逐字节不变（用户已生成的 seed.sql 的哈希会变成「差分あり」）。
- **md 的内容原样进生成物**：代码块的内容取未展开制表符的原始行（只按列去掉 fence 的缩进）；
  第一个 `##` 之前的数据块 / ```yaml、`-->` 后面同一行的内容、CSV 里 `,` 后未加引号的空格 —— 以前都是
  不报错地丢掉或改掉（数据被改、断言消失），现在一律报错。`TestDataIsNotSilentlyChanged` 守着。
- **写入是事务**（`apply_plans` → `_commit_files`）：同目录临时文件 → 退避既有文件后 `os.replace` →
  全部成功才删退避；任何一步失败都回滚（Windows 上 Excel 锁着期待值 CSV 是最常见的原因）。
  以前逐个覆盖，失败后新旧混在一起、记录的哈希对不上，之后 gen 一直判「手改」而卡住。
- `--check`（不写；全部「変更なし」且没有从 md 删掉的用例才 0）/ `--json`（1 行；`GenError.records` 带 line/case/message）。
- 普通 `gen` 也一样：有 差分あり / 対象外（没反映 md）的用例时退出码 1，并把它们从「次の手順」里去掉
  （以前返回 0 并照样列进 validate / run 的命令，旧 YAML 被当成按 md 验证过）。
- 改了生成器的输出格式，要重新生成 `demo/md_workflow` 的样例（`tests.test_casegen` 会提醒）：
  `cd demo/md_workflow && PYTHONPATH=../../src python -m autotest gen --md docs/testcases/order_import.md --config config/settings.yaml --overwrite`

## layout_txt 子系统

从 Excel 的布局定义生成 OCR 取込测试用的数据。**和用例/DB/证迹体系完全无关**，
改这边不会影响 autotest，反之亦然——但它住在 `src/autotest/` 下，所以同样受 3.6 兼容约束。

- **`layout_txt.py`** 是核心：读 Excel，按固定顺序生成 CSV 形式的 1 行 = 1 数据单位
  （FormID 情报 → 对象有无 → 之后 FieldID / OCR 文字识别结果 / 属性 flag / 座标 按对象项目数重复）。
  OCR 值由 Excel 的 **I/J/K 列（数据型 / IME / 最大位数）** 决定。
  **列位置和输出格式全部走命令引数**，就是为了不让「数据种别不同」变成复制一份代码。
  生成的是**测试数据的变体**，这是这个工具存在的理由：`--profile normal|max|over`（正常值/最大位数/超长）、
  `--date-mode`（和暦/西暦/跨年度 coverage 等）、`--error-patterns none|core|all`。
  加新变体时优先加 flag，不要新开一个模块。
- **`layout_tar.py`** 把正面/背面图像、识别结果 TXT、任意 CSV 打成一个 TAR。
  正面用后缀 `F`、背面用 `R`；正面 TXT 保持既有 FORM 生成内容，只有背面 TXT 作为可编辑字段暴露。
- **`layout_txt_gui.py`** 是独立 GUI（不是 autotest 那个 `gui.py`）。
- **`layout_txt_settings.py`**：GUI 设定**刻意存在应用文件夹之外**——
  `%APPDATA%\AUTO_TEST_BATCH\layout_txt_gui.json`（非 Windows 是 `~/.auto_test_batch/`），
  可用环境变量 `AUTOTEST_LAYOUT_GUI_CONFIG` 覆盖。原因：分发形态是 onedir EXE 文件夹，
  整包替换升级时写在里面的设定会被冲掉，而且那个目录可能只读。

## EXE 打包与 CI

layout_txt 要分发给没有 Python 的人，所以有一条独立的打包链。

```bash
build_layout_exe.bat --install    # 装 PyInstaller 4.10 后打包
build_layout_exe.bat              # 环境已备好时
#   可用 set LAYOUT_BUILD_PYTHON=C:\ProgramData\Anaconda3\python.exe 指定解释器
```

几个不写下来就会踩的点：

- **PyInstaller 钉死 4.10**（`requirements-build-py36.txt`），因为那是支持 Python 3.6 的最后一条线。
  BAT 会先检查解释器在 **3.6～3.10** 之间，不在就直接失败。
- **`pyinstaller_compat.py`** 打的是 Anaconda 5.2 特有的坑：它那版 Python 3.6 把
  `sysconfig._get_sysconfigdata_name` 的 `check_exists` 改成了必需参数，而 PyInstaller 4.10
  按上游签名无参调用。补丁**只在签名确实要求位置参数时才打**，不无条件覆盖。
- **产物是 onedir 不是单文件**：`dist/LayoutTxtGenerator/`。**要整个文件夹一起分发**，只拷 .exe 会跑不起来。
- **`layout_txt_exe.py --smoke-test`** 只 import 全部运行时依赖（tkinter / openpyxl / Pillow / 两个 layout 模块）
  而不开窗口。打包前后各跑一次——打包前验环境，打包后验产物。改依赖时记得同步这个 import 列表，
  否则缺库要等用户双击才发现。

CI 是 `.github/workflows/build-layout-exe.yml`（windows-2022 / Python **3.10.11**）：
按 path 过滤触发 → 装 `openpyxl==3.0.10` `Pillow==9.5.0` `pytest==7.4.4` →
**pytest** 跑 `test_layout_txt` / `test_layout_txt_exe` / `test_pyinstaller_compat` →
`build_layout_exe.bat --install` → 上传 artifact。

注意两件事：CI 用 3.10 构建，**真正保证 3.6 目标的是 `check_py36_compat.py` 而不是这条流水线**；
以及新增 layout 相关文件时要同步 workflow 的 `paths:` 列表，否则改了不触发构建。

## 改代码时最要紧的一条：偽 OK 是最重的缺陷

这个工具是判定别人对不对的裁判，所以「本来是 NG 却报 OK」比崩溃严重得多。全部代码和
`tests/test_regressions.py`（回帰テスト）都围绕这条展开。既有的防线，改动时不要削弱：

- `db.to_text()` **绝不整形值**（不四舍五入、不截位、不缩短）。整形后 `100.004` 和 `100.001` 会看起来一样。
  マスク也只在写进 Excel 时施加，比较用生值（`Table.mask_columns`）。（README 早期写过的
  `snapshot.format.decimal_places` 从未实现，已删除；snapshot 的未知键现在是硬错误。）
- 表示用的打切（`max_db_rows` / `max_lines_in_excel`）**不能进判定路径**。`truncated_from` 非空的 Table
  拿去比较会直接判 NG，而不是拿残缺数据比出个 OK。
- **初回作成ログは時刻フィルタしない**。実行前に存在しなかったファイルは全行が今回分。
  C# StreamWriter の UTF-8 BOM や時刻無しヘッダで先頭行を落とさない。ローテート時だけ時刻抽出する。
- **执行 0 件不是成功**。`load_cases` 在没有可执行用例时抛 `ConfigError`；`RunResult.verdict` 里
  「全 SKIP」也不算 OK。筛选条件和未执行件数会写进 Excel サマリ，避免「3 件全 OK」被误读成全量通过。
  `--include-disabled` 只给 validate / list（检查草稿），run / manual 拒绝它。
- **什么都不确认的判定不能出 OK**：`assert.log` 两个列表都空（非 manual）是 schema 错误；
  `execute.expected_exit_code` 曾被接受却没人读（以为确认了退出码），现在是带指引的错误。
- `manual: true` 的项即使自动比较一致也只给 `REVIEW`（要確認），永不自动转 OK，退出码走 3。
  子字段和 boolean 会严格校验；`review.py` 从 Excel 读取确认结果，自动 NG 不可覆盖，
  `finalize` 另存最终 Excel 和审计 JSON，同时保留原始自动判定。
- 同一 pattern 匹配到多个文件时判 NG 而不是取第一个——多出来的异常输出不能被静默忽略。
- **`mode: manual` 的用例整体封顶为 REVIEW**（`_force_review` + 永远置顶的 banner check）。
  `exit_code` 一律 SKIP——手动起动取不到，**且不接受用户自报**（无法验证的值不能发合格证）。
- **跳过的用例必须在证迹里留痕**：`RunResult.manual_pending` 会写进 Excel サマリ和控制台，
  并让总合判定保持 REVIEW / 退出码 3，防止「自动分全 OK」被读成「全部通过」。
  同理筛选执行会记录未执行件数。
- **比较不吞前后空白**：`compare._norm` 只把 CRLF / CR 归一成 LF，**不 strip**。曾经 strip 过，
  结果空白 1 个之差（固定长的填充缺失、右对齐数字栏的前导空白）全判成一致 —— PostgreSQL 的阴性对照测出来的。
  只差空白时 NG 的 detail 会注明（`_WHITESPACE_NOTE_*`），但判定绝不放宽；`tests/test_compare_strict.py` 守着。
- **「连上了」不等于「连对了」**：`dbcheck` 会 `SELECT DB_NAME(), @@SERVERNAME` 问 DB 本身，
  实际库和 `database.database` 不一致时**返回 1 而不是 0**。连进错误的库还发合格证，
  等于让整轮试验的证迹失效却无人察觉。GUI 顶部的 DB 栏默认是「未確認」（琥珀）而不是绿色，
  显示的是**实际连到的**服务器/库；改配置文件会把状态重置回未確認。
  这条链路的输出格式被 `gui._scan_line` 解析，改 `dbcheck` 的输出行会连带影响画面
  （`tests/test_dbcheck.py` 里有守这个契约的测试）。
- **结合证迹只能重建，不能往既有 xlsx 追加**：openpyxl 2.5（Anaconda 5.2 同梱）
  打开带图片的 xlsx 再保存会丢掉全部图片 = 静默销毁证迹。所以每个 case 跑完立刻
  `result_store.save_case_result` 落盘，`report` 从这些 JSON 重建新工作簿。
  **不要为了"简化"改回追加方式。** 另外 `report` 只重排版既定 verdict，不重算比较——
  证迹的结论不允许事后改变。
- **Spring Boot 的退出码不可信**：`System.exit(SpringApplication.exit(...))` 漏写、`.bat` 不回传
  ERRORLEVEL 时，作业 FAILED 也是 0。`job_status` 找不到完成行判 NG，不判 SKIP。
  起动行与完成行按（作业名, 参数）配对，按输出源分别计数取最大值（控制台和日志文件会重复出现同一行）；
  对象作业有没完成的执行 = NG，其他作业 = REVIEW。以前只收集完成行，一次完成就掩盖了后面中途停下的执行。

写新逻辑时的判据：**歧义一律往 NG / REVIEW 倒，绝不往 OK 倒**，并在注释里写清为什么。

## 破坏性操作的护栏

- `fsops.assert_safe_to_clear()` 是最后一道防线：拒绝文件系统根、用户 home 及其父、项目根、
  以及层级 ≤2 的路径（`C:\work`）。改动清空/删除逻辑时必须经过它。
- `preflight_case()` 在**任何破坏性动作之前**跑完所有静态检查（exe / jar / java 在不在、期待值文件全不全、
  clean_dirs 会不会连带删掉别的论理名……）。避免「清空了文件夹、重置了 DB，然后才发现 exe 不存在」。
  `validate` 命令与实际执行共用这一个函数——所以往 preflight 加检查会同时增强两边。
- teardown 和 `replace_files` 的还原都在 `finally` 里，中途异常也要把环境还回去。`--keep-env` 是唯一例外
  （调查用），此时会显式往证迹里记一条「没有还原」。
- `gen` 不覆盖手改过的文件、不删除 md 中消失的用例（只提示），写入前先完成全部校验。

## 加用例功能时的连带修改

所有 YAML（用例、settings、md 的 front matter / ```yaml 块 / `{…}` 选项）都用 `config.load_yaml_text`
（`UniqueKeyLoader`）读：**重复键是硬错误**。`yaml.safe_load` 会让后者静默覆盖前者，丢掉的检查项
在 schema 校验里看不出来。新增读取 YAML 的地方也必须用它（`<<` 合并键不算重复）。

用例 YAML 的未知键是**硬错误**（拼错了就报错退出，不静默忽略；下位 mapping —— snapshot.tables、
input_files、collect.files、teardown、sql 项 —— 也一样）。所以新增一个字段要同步：

1. `config.py` 的 `CASE_TOP_KEYS` / `SETUP_KEYS` / `EXECUTE_KEYS` / `COLLECT_KEYS` / `ASSERT_KEYS` 及各下位 `*_KEYS`
2. `orchestrator.preflight_case()` 里对应的静态检查
3. `README.md` 第 4 节的用例定义说明，以及 `docs/ai/case-schema.md`（`tests.test_docs_ai` 检查键是否都写进去了）
4. `templates/*.yaml`（4 个模板是新用户的实际入口，注释里没写的功能等于不存在）
5. 若 md 也要能写：`casegen.LABELS` / 数据块属性 + `docs/ai/testcase-format.md` + `tests/test_casegen.py`

**特别注意 orchestrator 读了但许可键里没有的字段**——那样的字段会在载入阶段被当成拼写错误拒绝，
按文档写的用例根本跑不起来。（`setup.db_lock` 曾经就是这种状态，现已补进 `SETUP_KEYS`。）
反过来，**许可键里有但没人读的字段**同样危险（`execute.expected_exit_code` 就是这样被删掉的）。

## 沟通与文体

代码注释、YAML 注释、CLI 输出、Excel 内容一律日文，README 用中文。注释的重点写「为什么这么做」
而不是「做了什么」——现有代码里大量注释记录的是某个偽 OK 事故或 Windows 上的坑，保持这个密度。
