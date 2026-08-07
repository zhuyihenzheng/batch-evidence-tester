# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

C# 制 batch（无界面 .exe）的自动测试 + 证跡（エビデンス）采集工具。目标运行环境是 **Windows + SQL Server + Anaconda 5.2**，但开发/验证可以在 macOS 上用 `--offline` 沙箱完成。
详细的使用说明、用例 YAML 全字段、Excel 成果物结构见 `README.md`（本项目的主文档）。

## 硬性约束

### Python 3.6 语法兼容

对象环境含 Anaconda 5.2（Python 3.6.5），**3.7+ 的语法/API 会在 import 时直接炸掉**。
所以：不用 dataclass（`models.py` 全是手写 `__init__`）、不用 `subprocess(capture_output=/text=)`、
不用 `Path.is_relative_to`（用 `fsops._is_relative_to`）、类型标注用 `typing.List` 而非 `list[...]`、
不用 f-string `=` 说明符 / 海象运算符 / `str.removeprefix`。

改动 `src/` 或 `demo/` 后必须过这一关：

```bash
python tools/check_py36_compat.py src/autotest demo
```

版本分支（`if sys.version_info >= (3, 12):`）是被允许的，检查器认识这种守卫。

### 不走 `pip install -e .`

社内封闭网络下构建隔离会去联网取 setuptools 而失败。所以 CLI 一律靠 `PYTHONPATH=src` 启动，
依赖只用 `requirements.txt`。`.bat` 脚本已内置这一行；手敲命令时要自己带上。

## 常用命令

```bash
# 单元测试（只用标准库 unittest，没有 pytest）
python -m unittest discover -s tests -v
python -m unittest tests.test_regressions.TestKeepEnvOption.test_teardown_skipped_and_reported  # 单个

# 3.6 兼容检查
python tools/check_py36_compat.py src/autotest demo

# 开发主循环：不碰真 DB / 真 exe 的沙箱全流程（macOS 上也能跑）
PYTHONPATH=src python -m autotest run --config config/settings.demo.yaml --offline
#   → demo/fake_batch.py 当 .exe，fixtures/ 当 DB，输出到 ./sandbox 与 ./output
#   → TC003 是「故意做成 NG 的样本」，所以这条命令正常情况下退出码就是 1

PYTHONPATH=src python -m autotest list                # 用例和 tag 一览
PYTHONPATH=src python -m autotest validate            # 设定 + 全用例 preflight（有问题退 1）
PYTHONPATH=src python -m autotest dbcheck             # 实际连一次 SQL Server（需 Windows/pyodbc）
PYTHONPATH=src python -m autotest run --dry-run       # 不启动 exe、不连 DB，只验流程

# 用例作成 / 手动实施 / 证迹结合
PYTHONPATH=src python -m autotest new --id TC010_x --template normal|error|env_missing|manual
PYTHONPATH=src python -m autotest copy --from TC001_normal --id TC011_x
PYTHONPATH=src python -m autotest manual --case TC008_manual_demo --phase before|after
PYTHONPATH=src python -m autotest report --run <run_id> --run <session> --out merged.xlsx
PYTHONPATH=src python -m autotest gui                 # 操作画面（需要 tkinter）
```

GUI 的自动化冒烟测试要 `root.withdraw()` + 只用 `root.update()` 驱动事件循环，
不要 `mainloop()`（会挂住）。

Windows 侧入口是 `.bat`（`setup_windows.bat` / `run_test.bat` / `run_demo.bat`），它们会自动选
`.venv` 或 PATH 上的 python 并设好 `PYTHONPATH`。macOS 上直接用上面的 CLI。

**退出码是对外契约**（タスクスケジューラ / CI 依赖它）：`0`=全 OK、`1`=有 NG、`2`=设定错误、`3`=有「要確認」。

## 架构要点

一个用例的完整生命周期集中在 `orchestrator.CaseRunner.run()` 一个方法里，顺序本身是设计（模块头部有注释说明为什么）：
preflight → setup（清空/SQL/投入）→ 実行前快照+文件夹撮影 → 记录日志字节偏移 → 跑 exe →
等日志写入落定 → 回收成果 → 実行後快照+撮影 → 渲染证据图 → 判定 → finally 里 teardown/还原。

- **`models.py` 是唯一的层间契约**。`excel.py` 只认识 `Table` / `CheckResult` / `CaseResult` / `RunResult`，
  不知道数据来自 DB、文件还是日志。加新证据类型时优先复用这几个类型，别让 excel.py 去认识新的来源。
- **文件夹论理名（path alias）是环境抽象**。用例只写 `input_dir`，物理路径只存在于 `settings.yaml` 的 `paths:`。
  `clean_dirs` / `remove_dirs` / `collect.files.dir` 只接受论理名——这是防误删的第一道闸。
- **配置文件自动选取**：存在 `config/settings.local.yaml` 就优先用它（已 gitignore，放各人的实路径和 DB 连接），
  否则用 `settings.yaml`（仓库维护的模板）。所以不要在 `run_test.bat` 里写死 `--config`。
- **offline 模式**：`fixtures/<case_id>/<before|after>_<TABLE>.csv` 顶替 DB 快照，setup SQL 不执行
  （所以靠 SQL 注入故障的用例在 offline 下无效，它们 `enabled: false`）。
- **期待值基线流程**：跑通一次 → 从 `output/<run_id>/artifacts/` 和 Excel 的「実行後」表取实际值 →
  人工确认 → 放进 `expected/<case_id>/` 当基线。
- **`mode: manual` 把生命周期劈成两个进程**：`run_manual_before` / `run_manual_after`
  复用 `run()` 的同一批私有方法（生命周期知识只有一份）。跨进程状态存
  `output/manual_<id>_<ts>/session.json`，字段是固定 schema，`format_version` 不匹配直接报错。
  刻意**不**持久化 `replace_files` 退避信息和 `db_lock`——它们依赖 `run()` 的 `finally` 还原，
  两个进程之间没有 finally，所以 preflight 直接禁止在 manual 用例里使用。
  `base_date` 必须从 session 恢复（用当天日期会让跨日的 `{date}` 静默错配）。
- **GUI 是 CLI 之上的一层**：`gui.py` 一律 `subprocess` 调 `python -m autotest ...`，
  绝不在进程内调 `CaseRunner`。原因：长跑/pyodbc 崩溃不能拖死界面；tkinter 非线程安全；
  编排入口保持一条（150+ 测试守的是 CLI 这条路）。进度靠守护线程读 stdout → `queue.Queue`
  → `root.after()` 排水，Tk 组件只在主线程碰。
- **用例发现是递归的**：`cases/受注/取込/TC001.yaml` 自动带上 tag `[受注, 取込]`；
  `cases/TC001/` 这种「与定义文件同名的目录」被视为该用例的资材目录，其中的 YAML 不会被当成用例。
- **一个 YAML = 一个用例**，没有第二种格式。曾经做过「一个文件多用例（`defaults:` + `cases:`）」，
  已删除——它和 `autotest copy` / GUI 的複製按钮互斥，而且深合并的「列表整个替换」会让
  case 级的 `assert.db` 静默顶掉 defaults 继承的检查（断言变少、零警告，正是偽 OK 方向）。
  要共享共通部分，走**显式的按文件引用**（如 `sql: {file:}`），不要做隐式合并。

## 改代码时最要紧的一条：偽 OK 是最重的缺陷

这个工具是判定别人对不对的裁判，所以「本来是 NG 却报 OK」比崩溃严重得多。全部代码和
`tests/test_regressions.py`（回帰テスト）都围绕这条展开。既有的防线，改动时不要削弱：

- `db.to_text()` **绝不整形值**（不四舍五入、不截位、不缩短）。整形后 `100.004` 和 `100.001` 会看起来一样。
  マスク也只在写进 Excel 时施加，比较用生值（`Table.mask_columns`）。
- 表示用的打切（`max_db_rows` / `max_lines_in_excel`）**不能进判定路径**。`truncated_from` 非空的 Table
  拿去比较会直接判 NG，而不是拿残缺数据比出个 OK。
- **执行 0 件不是成功**。`load_cases` 在没有可执行用例时抛 `ConfigError`；`RunResult.verdict` 里
  「全 SKIP」也不算 OK。筛选条件和未执行件数会写进 Excel サマリ，避免「3 件全 OK」被误读成全量通过。
- `manual: true` 的项即使自动比较一致也只给 `REVIEW`（要確認），永不自动转 OK，退出码走 3。
- 同一 pattern 匹配到多个文件时判 NG 而不是取第一个——多出来的异常输出不能被静默忽略。
- **`mode: manual` 的用例整体封顶为 REVIEW**（`_force_review` + 永远置顶的 banner check）。
  `exit_code` 一律 SKIP——手动起动取不到，**且不接受用户自报**（无法验证的值不能发合格证）。
- **跳过的用例必须在证迹里留痕**：`RunResult.manual_pending` 会写进 Excel サマリ和控制台，
  防止「自动分全 OK」被读成「全部通过」。同理筛选执行会记录未执行件数。
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

写新逻辑时的判据：**歧义一律往 NG / REVIEW 倒，绝不往 OK 倒**，并在注释里写清为什么。

## 破坏性操作的护栏

- `fsops.assert_safe_to_clear()` 是最后一道防线：拒绝文件系统根、用户 home 及其父、项目根、
  以及层级 ≤2 的路径（`C:\work`）。改动清空/删除逻辑时必须经过它。
- `preflight_case()` 在**任何破坏性动作之前**跑完所有静态检查（exe 在不在、期待值文件全不全、
  clean_dirs 会不会连带删掉别的论理名……）。避免「清空了文件夹、重置了 DB，然后才发现 exe 不存在」。
  `validate` 命令与实际执行共用这一个函数——所以往 preflight 加检查会同时增强两边。
- teardown 和 `replace_files` 的还原都在 `finally` 里，中途异常也要把环境还回去。`--keep-env` 是唯一例外
  （调查用），此时会显式往证迹里记一条「没有还原」。

## 加用例功能时的连带修改

用例 YAML 的未知键是**硬错误**（拼错了就报错退出，不静默忽略）。所以新增一个字段要同步三处：

1. `config.py` 的 `CASE_TOP_KEYS` / `SETUP_KEYS` / `EXECUTE_KEYS` / `COLLECT_KEYS` / `ASSERT_KEYS`
2. `orchestrator.preflight_case()` 里对应的静态检查
3. `README.md` 第 4 节的用例定义说明
4. `templates/*.yaml`（4 个模板是新用户的实际入口，注释里没写的功能等于不存在）

**特别注意 orchestrator 读了但许可键里没有的字段**——那样的字段会在载入阶段被当成拼写错误拒绝，
按文档写的用例根本跑不起来。（`setup.db_lock` 曾经就是这种状态，现已补进 `SETUP_KEYS`。）

## 沟通与文体

代码注释、YAML 注释、CLI 输出、Excel 内容一律日文，README 用中文。注释的重点写「为什么这么做」
而不是「做了什么」——现有代码里大量注释记录的是某个偽 OK 事故或 Windows 上的坑，保持这个密度。
