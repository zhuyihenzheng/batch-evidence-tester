# AUTO_TEST_BATCH — batch(.exe) 自动测试工具

C# 制 batch（无界面 .exe）的自动化测试工具。运行环境：**Windows + SQL Server**。

一条命令完成：**准备输入 → 跑 exe → 采集 DB/日志/文件 → 与期待值自动比对 → 生成 Excel 证据簿（エビデンス）**。

---

## 1. 方案总览

```
                 ┌──────────────── 1 个用例的生命周期 ────────────────┐
  cases/*.yaml   │                                                  │
   （用例定义）  │  ① 清空文件夹 + 执行 setup SQL + 投入测试数据      │
        │        │  ② DB 快照【実行前】 + 文件夹一览截图【実行前】    │
        ▼        │  ③ 记录日志文件字节偏移                           │
  ┌──────────┐   │  ④ 启动 batch.exe（记录 exit code / 耗时 / 控制台）│
  │ autotest │──▶│  ⑤ 切出本次执行新增的日志行                       │
  │  runner  │   │  ⑥ 回收输出/処理済/エラー文件夹的文件并保全        │
  └──────────┘   │  ⑦ DB 快照【実行後】 + 文件夹一览截图【実行後】    │
        │        │  ⑧ 渲染证据图（文件夹一览 / 取回文件的内容）       │
        │        │  ⑨ 与 expected/ 比对 → 每项 OK/NG                 │
        ▼        └──────────────────────────────────────────────────┘
  output\TestEvidence_<run_id>.xlsx   ← 成果物
  output\<run_id>\evidence\<case>\*.png
  output\<run_id>\artifacts\<case>\   ← 回收的实际文件（原样保全）
```

### 什么进 Excel 单元格，什么进图片

| 采集对象 | 形式 | 理由 |
|---|---|---|
| **DB 数据** | Excel 原生表格 | 可筛选、可搜索、可再比对；执行前后变化的单元格标黄 |
| **日志** | Excel 原生表格（行号 + 本文） | 可搜索可复制；ERROR 行标红、WARN 标黄、期待关键字标蓝 |
| **输入/输出文件夹一览** | **图片** | 「文件确实被移到処理済了」这种事实，一览画面最直观 |
| **取回文件的内容** | **图片** | 附 SHA-256 和字节数，作为「当时确实是这个内容」的证据 |
| 控制台输出 | Excel 原生文本 | — |

### 三个关键设计决策

| 决策 | 做法 | 理由 |
|---|---|---|
| **证据图用代码渲染**（非屏幕截图） | PIL 画成 Explorer 窗口的样子 | 不需要交互式登录会话 → 可挂タスクスケジューラ无人值守跑；同输入同输出，可复现；不受分辨率和窗口位置影响 |
| **日志只取增量** | 执行前记字节偏移，执行后只读新增部分 | 避免把历史日志全贴进证据。文件被 rotate 时降级为「时间戳过滤」，再降级为「全文」 |
| **文件夹用论理名** | 用例只写 `input_dir`，物理路径只在 settings 里 | 换环境只改一个文件；且 `clean_dirs` 只接受论理名，防止设定失误清掉无关目录 |

> 如果确实需要真实的 Explorer 窗口截图，把 `evidence.mode` 改成 `screen` 或 `both`，
> 再 `pip install mss pywin32` 即可。截图失败会自动回退到渲染方式，不会中断测试。

---

## 2. 在 VSCode 里装环境（Windows）

### 环境要求

**Python 3.6 以上**。代码刻意做了 3.6 兼容，所以 **Anaconda 5.2（Python 3.6.5）可以直接用**，
不需要新建环境。而且 Anaconda 5.2 已经自带了本工具需要的 openpyxl / Pillow / PyYAML，
**大概率一个包都不用装**。

| 库 | 用途 | Anaconda 5.2 自带 | 本工具最低要求 |
|---|---|---|---|
| openpyxl | 生成 Excel 成果物 | 2.5.3 | 2.5 |
| Pillow | 渲染证据图 | 5.1.0 | 5.1 |
| PyYAML | 读配置和用例 | 3.12 | 3.12 |
| pyodbc | 连 SQL Server（`--offline` 时不需要） | 通常自带 | 4.0 |

---

### 步骤 1：装 VSCode 的 Python 扩展

VSCode 本身不带 Python 支持，要先装扩展：

1. 按 `Ctrl + Shift + X` 打开扩展面板（或点左侧栏的方块图标）
2. 搜索框输入 `Python`
3. 选**发布者是 Microsoft** 的那个（扩展 ID：`ms-python.python`），点 **Install**

> 命令行装也行：`code --install-extension ms-python.python`

顺带推荐 `Pylance`（微软出的补全和跳转），非必需。

### 步骤 2：选择 Python 解释器

**这步最关键**，决定 VSCode 用哪个 Python 跑代码。

1. `Ctrl + Shift + P` 打开命令面板
2. 输入 `Python: Select Interpreter`，回车
3. 列表里选 Anaconda 那个，路径类似：
   ```
   C:\ProgramData\Anaconda3\python.exe          （全机器安装）
   C:\Users\<用户名>\Anaconda3\python.exe        （单用户安装）
   ```
   列表里没有就选 `Enter interpreter path...` 手动填。

选完后打开任意 `.py` 文件，**右下角状态栏**会显示当前解释器和版本，点它可随时切换。

### 步骤 3：打开终端并自检

按 `` Ctrl + ` ``（反引号，Tab 键上面那个）打开集成终端。

> **Anaconda 5.2 注意**：它自带的 conda 是 4.5 版，还没有 `conda init`，
> 所以 **PowerShell 里 `conda activate` 会失败**。二选一：
> - 终端右上角下拉 → **Select Default Profile** → 选 **Command Prompt**，然后用 `activate base`
> - 或从开始菜单开 **Anaconda Prompt**，`cd` 到项目目录操作

跑自检，**这步会明确告诉你缺不缺东西**：

```cmd
python check_env.py
```

输出示例：

```
Python      : 3.6.5  (C:\ProgramData\Anaconda3\python.exe)
OS          : Windows 10
conda 環境  : base   prefix=C:\ProgramData\Anaconda3
conda       : conda 4.5.11
Anaconda 版 : 5.2.0
              ※ Python 3.6 系ですが、本体は 3.6 対応済みのため動作します

必須パッケージ:
  [OK] openpyxl   2.5.3
  [OK] Pillow     5.1.0
  [OK] PyYAML     3.12

任意パッケージ:
  [OK] pyodbc     4.0.23

 結論: この環境で実行できます。
```

**全是 `[OK]` 就什么都不用装**，直接跳到步骤 5。

### 步骤 4：装缺的包（只在自检报 `[NG]` 时才需要）

Anaconda 环境**优先用 conda**，和 pip 混用容易把环境搞乱：

```cmd
conda install openpyxl pillow pyyaml pyodbc
```

conda 装不了再退回 pip：

```cmd
python -m pip install -r requirements.txt
```

> ⚠️ **两个企业环境常见的坑**
>
> 1. **Anaconda 商用许可**：2024 年起 Anaconda 官方频道（`defaults`）对达到一定规模的
>    组织商用需要付费许可。装之前跟公司确认，或改用 conda-forge 频道：
>    `conda install -c conda-forge openpyxl pillow pyyaml`
> 2. **封闭内网连不上外网**：先在能联网的机器上
>    `pip download -r requirements.txt -d packages\` 下载 whl，
>    拷进来后 `pip install --no-index --find-links=packages -r requirements.txt`

装完再跑一次 `python check_env.py` 确认。

### 步骤 5：连真实 DB 还需要两件事

只跑 `--offline` 沙箱验证的话可以跳过。

1. **ODBC 驱动**：装 [Microsoft ODBC Driver for SQL Server](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server)。
   pyodbc 只是个壳，真正连 SQL Server 靠这个驱动。装的版本要和 `settings.yaml` 的
   `database.driver` 对上（老机器上常见的是 `ODBC Driver 17 for SQL Server`）。
   查本机已装的驱动：
   ```cmd
   powershell "Get-OdbcDriver | Select-Object -Unique Name"
   ```
2. **DB 密码放环境变量**，不写进 YAML：
   ```cmd
   setx AUTOTEST_DB_PASSWORD "パスワード"
   ```
   设完要**重开终端**才生效（`setx` 只影响之后新开的进程）。
   用 Windows 认证的话把 `database.auth` 改成 `windows`，这步可跳过。

### 步骤 6：空跑一遍看成果物

```cmd
run_demo.bat
```

用 `demo\fake_batch.py` 代替真 exe、用 `fixtures\` 代替 DB、输出到 `.\sandbox\`，
**完全不碰真实环境**，产出完整的 `output\TestEvidence_*.xlsx`。
样例里 TC001/TC002 是 OK、TC003 是**故意做成 NG 的样本**，可以直接看到差分表和 NG 上色。

> `.bat` 会自动判断：有 `.venv` 就用 `.venv`，没有就用 PATH 上的 `python`（即 Anaconda 的）。
> 所以 **Anaconda 用户不需要跑 `setup_windows.bat`**。

### 在 VSCode 里直接跑（不用 .bat）

终端里敲：

```cmd
set PYTHONPATH=src
python -m autotest validate
python -m autotest run --config config\settings.demo.yaml --offline
```

想用 F5 调试，建个 `.vscode\launch.json`：

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "autotest run (sandbox)",
      "type": "python",
      "request": "launch",
      "module": "autotest",
      "args": ["run", "--config", "config/settings.demo.yaml", "--offline"],
      "cwd": "${workspaceFolder}",
      "env": { "PYTHONPATH": "${workspaceFolder}/src" },
      "console": "integratedTerminal"
    }
  ]
}
```

### 常见问题

| 现象 | 原因和处理 |
|---|---|
| `python` 提示不是内部或外部命令 | Anaconda 没加进 PATH。用 **Anaconda Prompt**，或填全路径 `C:\ProgramData\Anaconda3\python.exe` |
| PowerShell 里 `conda activate` 报错 | Anaconda 5.2 的 conda 4.5 不支持。换 **Command Prompt** 终端，用 `activate base` |
| `ModuleNotFoundError: No module named 'autotest'` | 没设 `PYTHONPATH=src`，或没在项目根目录下执行 |
| 终端里日文乱码 | 先敲 `chcp 65001` 切到 UTF-8 |
| 右下角解释器和终端里的 `python` 不一致 | VSCode 的解释器设置只影响运行/调试，不改终端 PATH。以 `check_env.py` 输出的 `sys.executable` 为准 |

---

## 3. 配置

改 `config\settings.yaml`：

```yaml
batch:
  exe_path: "C:/app/batch/OrderBatch.exe"   # ← 被测 exe（正斜杠即可）
  working_dir: "C:/app/batch"
  timeout_sec: 600
  console_encoding: "cp932"                 # ← 日文 Windows 控制台编码

paths:                                       # ★ 输入输出文件夹设定
  input_dir:     "C:/app/batch/in"          # 投入
  processed_dir: "C:/app/batch/processed"   # 処理済
  error_dir:     "C:/app/batch/error"       # エラー
  output_dir:    "C:/app/batch/out"         # 出力
  log_dir:       "C:/app/batch/log"         # ログ

folder_evidence:
  targets: [input_dir, processed_dir, error_dir, output_dir]   # 要拍一览图的文件夹

database:
  server: "SQLSRV01,1433"
  database: "BATCH_DB"
  auth: "sql"                # Windows 认证填 windows
  user: "sa"
  password_env: "AUTOTEST_DB_PASSWORD"
```

文件夹要增删，只改 `paths` 和 `folder_evidence.targets` 两处即可，用例不用动。

改完确认：

```cmd
.venv\Scripts\python -m autotest validate
```

会逐项打印每个文件夹是否存在、exe 是否找得到。全绿之后再跑正式测试。

---

## 4. 用例定义

一个 YAML 一个用例，五个区块对应生命周期：

```yaml
id: TC001_normal
name: "正常系：受注ファイル取込（明細2件）"
tags: [正常系, 単体]

setup:                                    # ① 前处理
  clean_dirs: [input_dir, processed_dir, error_dir, output_dir]
  sql:
    - file: sql/TC001_setup.sql           # 支持 GO 分隔；也可写 inline
  input_files:
    - src: "input/ORDER_20260803.csv"     # 相对 cases\<id>\
      dest_dir: input_dir

snapshot:                                 # ② ⑦ DB 快照（执行前后各取一次）
  tables:
    - name: T_ORDER
      sql: "SELECT ORDER_ID, ORDER_NO, AMOUNT, STATUS, UPDATED_AT FROM T_ORDER ORDER BY ORDER_ID"
      format: {decimal_places: {AMOUNT: 0}}
      mask: [CUSTOMER_TEL]                # 个人信息列 → ***MASKED***

execute:                                  # ④ 执行
  args: ["--mode", "daily", "--date", "20260803"]

collect:                                  # ⑥ 回收
  files:
    - dir: output_dir
      pattern: "RESULT_*.csv"
      preview: true                       # 中身也渲染成证据图

assert:                                   # ⑨ 判定
  exit_code: 0
  db:
    - table: T_ORDER
      expected: "expected/TC001_normal/db_T_ORDER.csv"
      key: [ORDER_ID]                     # 按主键对齐记录后逐单元格比
      ignore_columns: [UPDATED_AT]        # 依赖执行时刻的列排除
  files:
    - name: "結果ファイル"
      actual: {dir: output_dir, pattern: "RESULT_*.csv"}
      expected: "expected/TC001_normal/RESULT_20260803.csv"
      encoding: cp932
      ignore_line_patterns: ['^#作成日時']  # 行内含时刻等易变内容可正则排除
    - name: "処理済フォルダへの移動"
      exists: {dir: processed_dir, pattern: "ORDER_*.csv", count: 1}
    - name: "エラーフォルダが空であること"
      exists: {dir: error_dir, pattern: "*", count: 0}
  log:
    must_contain: ["処理開始", "取込件数=2", "処理正常終了"]
    must_not_contain: ["[ERROR]", "Exception", "異常終了"]
```

**期待值文件怎么来**：第一次跑通后，从 `output\<run_id>\artifacts\` 取实际输出、
从 Excel 的「実行後」DB 表复制成 CSV，人工确认无误后放进 `expected\` 当基线。
之后每次回归就是全自动的。

### 不同输入文件怎么组织

**一个用例 = 一组输入**。每个用例有自己的资材文件夹，投入文件放里面：

```
cases\
├─ TC001_normal.yaml              用例定义
├─ TC001_normal\input\            ← 这个用例投入的文件
│    └─ ORDER_20260803.csv
├─ TC002_error.yaml
└─ TC002_error\input\             ← 另一个用例的文件（内容不同）
     └─ ORDER_20260804.csv
```

用例里 `src` 写相对 `cases\<用例ID>\` 的路径：

```yaml
setup:
  input_files:
    - src: "input/ORDER_20260803.csv"
      dest_dir: input_dir              # 投到哪个文件夹（论理名）
    - src: "input/MASTER.csv"          # 一次投多个文件也行
      dest_dir: input_dir
      rename: "MASTER_LATEST.csv"      # 需要改名时用
```

想测同一个 batch 的不同输入（正常/异常/边界/空文件/大容量），就是**复制一份用例 YAML +
换掉 `input\` 里的文件 + 改期待值**，三步。文件夹路径、DB 连接这些都不用碰。

### 多个 batch 怎么组织

`settings.yaml` 里用 `batches:` 定义多个 exe，用例里用 `execute.batch` 选：

```yaml
# config/settings.yaml
batch:                                    # ← 既定 batch（不写 execute.batch 时用这个）
  exe_path: "C:/app/batch/OrderBatch.exe"
  working_dir: "C:/app/batch"
  console_encoding: "cp932"
  timeout_sec: 600

batches:                                  # ← 名前付き定义，只写和既定的差异
  invoice:
    exe_path: "C:/app/invoice/InvoiceBatch.exe"
    working_dir: "C:/app/invoice"
    log_dir: inv_log_dir                  # 这个 batch 的日志在别处
  nightly:
    exe_path: "C:/app/batch/NightlyBatch.exe"
    timeout_sec: 3600                     # 只覆盖超时，其余继承 batch:

paths:
  input_dir:      "C:/app/batch/in"       # 既定 batch 的文件夹
  processed_dir:  "C:/app/batch/processed"
  output_dir:     "C:/app/batch/out"
  log_dir:        "C:/app/batch/log"
  inv_input_dir:  "C:/app/invoice/in"     # invoice batch 的文件夹（另一套）
  inv_output_dir: "C:/app/invoice/out"
  inv_log_dir:    "C:/app/invoice/log"
```

```yaml
# cases/TC004_invoice.yaml
setup:
  clean_dirs: [inv_input_dir, inv_output_dir]
  input_files:
    - src: "input/INVOICE_20260803.csv"
      dest_dir: inv_input_dir

execute:
  batch: invoice                          # ← 选哪个 batch
  args: ["--mode", "monthly"]

collect:
  folder_evidence: [inv_input_dir, inv_output_dir]   # ← 截图对象也换成这套文件夹
  files:
    - dir: inv_output_dir
      pattern: "RESULT_*.csv"
      preview: true
```

三个要点：

| 要素 | 机制 |
|---|---|
| **exe** | `batches.<名前>` 是 `batch:` 的**差分**，共通项目（编码、超时）不用重写 |
| **文件夹** | `paths` 里加一套新的论理名即可，用例用名字引用 |
| **日志** | `batches.<名前>.log_dir` 指定该 batch 的日志目录，不写则用 `paths.log_dir` |
| **截图对象** | `collect.folder_evidence` 按用例覆盖，不写则用 `folder_evidence.targets` |

样例见 `cases\TC004_other_batch.yaml`（跑 `run_demo.bat` 能看到它和默认 batch 在同一份
Excel 里各占一个 sheet）。

> 什么时候该拆成两份 `settings.yaml` 而不是用 `batches:`？
> **环境不同就拆**（结合测试环境 vs 本番相当环境，DB 连接不同），用 `--config` 切换；
> **同环境下的不同 batch 就用 `batches:`**，一次执行出一份 Excel。

---

## 5. Excel 成果物结构

| Sheet | 内容 |
|---|---|
| **サマリ** | 实行环境、实施者、总合判定、各用例一览（判定/NG 件数/exit code/耗时），带筛选和跳转链接 |
| **<用例ID>** | 1. 実行情報（命令行、开始终了时刻、exit code）<br>2. 判定明細（逐项 OK/NG，绿/红上色）<br>3. DB スナップショット 実行前/実行後（原生表格，变化单元格标黄）<br>4. 実行ログ（原生表格，行号+本文，ERROR 行标红）<br>5. 期待値との差分（NG 时才出，1 行 1 处相违）<br>6. 証跡画像（入出力フォルダ一览 実行前/実行後、取回文件内容）<br>7. 標準出力 / 標準エラー |

差分表格式（DB）：`区分 | キー | 列名 | 期待値 | 実績値`，区分有 `相違` / `期待のみ` / `実績のみ`。
文件差分是行单位：`行 | 区分 | 期待値 | 実績値`。

无对应数据的章节会自动跳过，编号连续不跳号。

---

## 6. 运行

### 选择要跑哪些用例

```cmd
python -m autotest list                 :: 先看有哪些用例和标签

run_test.bat                            :: 全部（enabled: true のもの）
run_test.bat --case TC001_normal        :: 单个
run_test.bat --case TC001 --case TC005  :: 多个（--case 重复即可）
run_test.bat --tag 異常系                :: 按标签
run_test.bat --tag 正常系 --tag 環境不備  :: 多标签 = 任一命中（OR）
run_test.bat --case TC001 --tag 異常系    :: 同时用 = 两者都满足（AND）
```

用例的标签在 YAML 里定义：

```yaml
tags: [異常系, 環境不備]
```

打错 ID 或标签会**报配置错误并退出码 2**，不会静默跑 0 件。

```cmd
:: --- 接真实环境前的确认 ---
python -m autotest dbcheck            :: SQL Server 能不能连上（列出已装 ODBC 驱动 + 实际连一次）
python -m autotest validate           :: 配置和用例的完整性（有问题退出码 1）
python -m autotest run --dry-run      :: 不碰 DB、不启动 exe，只验流程

:: --- 正式执行 ---
run_test.bat                          :: 全用例
run_test.bat --case TC001_normal      :: 指定用例（可多次指定）
run_test.bat --tag 正常系              :: 按标签
run_test.bat --dry-run                :: 不启动 exe、不动文件夹，只验流程
run_demo.bat                          :: 沙箱空跑

:: 直接用 CLI 也可以（.bat 里已自动设好，手敲时需要这一行）
set PYTHONPATH=src
.venv\Scripts\python -m autotest run --config config\settings.yaml
.venv\Scripts\python -m autotest validate
.venv\Scripts\python -m autotest list
```

> 没走 `pip install -e .` 是刻意的：那条路径会因为构建隔离去联网取 setuptools，
> 封闭的社内网络容易失败。`requirements.txt` + `PYTHONPATH` 更稳。
> 若网络允许，`pip install -e .` 后可直接用 `autotest` 命令，两种方式都支持。

每次运行都会写 `output\<run_id>\run.log`，逐工程记录并立即刷盘。卡住时另开窗口
`Get-Content output\<run_id>\run.log -Wait -Tail 20` 就能看到停在哪一步。
加 `--verbose` 可以直接在屏幕上看。

**退出码**：全 OK = `0`，有 NG = `1`，设定错误 = `2`
→ 可直接接 タスクスケジューラ / Jenkins / GitLab CI。
`run_test.bat` 从调度器启动时不会 `pause`，手动双击时才停留等按键。

某个用例抛异常不会中断整轮，会记为该用例的 `ERROR` 并继续跑后续用例。

---

## 7. 目录结构

```
AUTO_TEST_BATCH\
├─ setup_windows.bat      初次安装
├─ run_test.bat           正式执行
├─ run_demo.bat           沙箱空跑
├─ config\
│   ├─ settings.yaml      ★ 全局设定
│   └─ settings.demo.yaml 沙箱用
├─ cases\
│   ├─ TC001_normal.yaml  用例定义
│   └─ TC001_normal\input\ 该用例的投入文件
├─ expected\<case_id>\    期待值
├─ fixtures\<case_id>\    offline 模式下代替 DB 的 CSV
├─ sql\                   setup SQL（支持 GO 分隔）
├─ demo\fake_batch.py     模拟 batch
├─ src\autotest\          工具本体
└─ output\                成果物
```

| 模块 | 职责 |
|---|---|
| `cli.py` | 命令行入口（`run` / `validate` / `list`） |
| `config.py` | 读取校验 settings 与用例；文件夹论理名 → 物理路径解析 |
| `orchestrator.py` | 单个用例的生命周期编排（上图 ①〜⑨） |
| `db.py` | SQL Server 连接、快照取得、值格式化；offline 模式的 CSV 替身 |
| `fsops.py` | 文件夹清空 / 投入 / 回收 / 一览取得 |
| `runner.py` | 启动 .exe，取 exit code、stdout/stderr、耗时 |
| `logs.py` | 日志增量切出、行分类、关键字断言 |
| `compare.py` | 期待值比对，生成差分表 |
| `render.py` | 证据图渲染 |
| `screenshot.py` | 真实屏幕截图（可选） |
| `excel.py` | Excel 成果物生成 |

---

## 8. Windows 实施上的注意点

| 项目 | 说明 |
|---|---|
| **文字编码** | .exe 控制台输出多为 cp932，日志文件可能是 UTF-8 或 cp932。`batch.console_encoding` 和 `log.encoding` + `encoding_fallbacks` 分别指定，读取失败依次降级不会崩。控制台显示乱码时先 `chcp 65001` |
| **字体** | 证据图用 `C:\Windows\Fonts\meiryo.ttc`（`evidence.font_candidates` 第一顺位），日文 Windows 标配 |
| **文件被占用** | 清空文件夹时若 batch、Explorer、编辑器或杀软仍握着句柄会失败，报错会指出具体文件。只读属性会自动解除后重试 |
| **路径** | YAML 里用正斜杠 `C:/app/batch` 即可，Python 在 Windows 下同样识别。相对路径一律以项目根目录为基准，不受启动目录影响 |
| **时钟偏差** | DB 服务器和测试机时刻不一致时，日志按时间戳切片会漏行。默认走字节偏移方式规避 |
| **并发** | 用例串行执行。多个用例共用同一组文件夹和 DB 表，并行会互相污染 |
| **大数据量** | DB 快照超过 `excel.max_db_rows`（默认 500 行）、日志超过 `log.max_lines_in_excel`（默认 500 行）会中略并注明总数 |
| **个人信息** | `snapshot.tables[].mask` 列会替换成 `***MASKED***` 后再写进 Excel，证据簿可直接共享 |

---

## 9. 后续可扩展方向

- **套用公司 Excel 模板**：现在是从零生成。若有固定的エビデンス様式，改 `excel.py` 的
  `build_workbook` 为「打开模板 → 往指定位置填」即可，其余模块不用动。
- **跨表判定**：目前 DB 判定是单表对单 CSV。需要跨表校验时，在 `snapshot.tables` 里
  直接写 JOIN 的 SQL，把结果当成一张虚拟表比对。
- **性能回归**：`ExecutionInfo.elapsed_sec` 已记录，加个阈值断言即可。
