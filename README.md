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

### 操作画面（不想记命令的话用这个）

```cmd
run_gui.bat
```

一个窗口里能做完日常的事：

| 区域 | 能做什么 |
|---|---|
| 用例一览 | 显示 ID / 区分（tag）/ 实行方式（自动·手动）/ 用例名。点最左列的 ☑ 勾选 |
| 上方 tag 下拉 | 按 tag 过滤一览；「タグで実行」直接按 tag 跑 |
| 「全件実行」「選択したケースを実行」 | 全部跑 / 只跑勾选的 |
| 「新規ケース作成」「選択ケースを複製」 | 从模板建新用例 / 复制既有用例（见下节） |
| 「手動：実行前を採取」「手動：実行後を採取」 | 手动实施用例的两段式证迹采集（见 6.4） |
| 下方黑色区域 | 实时显示执行进度（就是命令行的输出） |
| 「中止」 | 停止执行（会确认；证迹会不完整） |
| 「証跡Excelを開く」 | 跑完直接打开成果物 |

底部状态栏用颜色显示总合判定：绿=OK / 红=NG / 黄=要確認。

### ★ 接続先 DB 一直显示在画面上

连错 DB 是最难发现的一类错误——结果看着完全正常，其实测的是别的库，而这份证迹会
被当成有效的提出物。所以画面顶部**常驻**一条 DB 栏：

```
環境: 結合テスト環境    接続先: SQLSRV01,1433 / BATCH_DB  （SQL Server認証: sa）    ● 未確認（実行前に確認）  [DB接続確認]
```

| 状态 | 显示 | 底色 |
|---|---|---|
| 还没确认过 | `● 未確認（実行前に確認）` | 琥珀 |
| 确认过、和设定一致 | `● 確認済: <实际服务器> / <实际DB>` | 绿 |
| **连上了，但连的不是设定的库** | `▲ 接続先が設定と相違: ...` + 弹窗警告 | 红 |
| 连不上 | `✕ 接続できません` | 红 |
| 勾了オフライン | `オフライン（DB は使いません）` | 灰 |

三个要点：

1. **默认是「未確認」不是绿色**。没确认过的东西不给它看起来没问题。
2. **显示的是「实际连到的」，不是「设定里写的」**。`DB接続確認` 按钮跑的是
   `autotest dbcheck`，它会 `SELECT DB_NAME(), @@SERVERNAME` 问 DB 本身。
   只把设定念一遍是发现不了连错的——你设定写对了，登录名的既定数据库把你带到别处，
   这种情况只有问 DB 才知道。
3. **改了配置文件，确认状态自动回到「未確認」**。旧的确认结果对新的接続先不作数。

> `dbcheck` 的退出码：连上且库对 = `0`，**连上但库不对 = `1`**，连不上 = `1`。
> 「连上了」本身不算合格——接进错误的库还返回 0 的话，就等于给错误的试验发了张事前确认合格证。

> **画面只是操作台**。执行本身还是调 `python -m autotest run`，所以画面里跑和命令行跑
> 结果完全一样。无人值守（タスクスケジューラ / CI）继续用 `run_test.bat`。
>
> ⚠️ **但两者不能同时跑**。所有用例共用同一组文件夹和 DB 表，并行会互相污染
> （见第 8 节）。画面本身会禁止重复启动执行，但它管不到另一个窗口或调度器起的
> `run_test.bat`。挂了定时任务的话，注意避开那个时间段。
>
> 画面打不开时：`python check_env.py` 看 tkinter 那行。tkinter 是标准库，
> Anaconda 一定自带；万一没有，命令行的全部功能不受影响。

### 追加用例：用模板，不要手写 YAML

```cmd
python -m autotest new --id TC010_受注取込_0件            :: 从模板建
python -m autotest new --id TC011_金額不正 --template error
python -m autotest copy --from TC001_normal --id TC012_明細100件  :: 复制既有用例
```

`new` 会一次建好三样东西，并打印接下来该做什么：

```
cases\TC010_受注取込_0件.yaml     ← 用例定义（带注释的填空模板）
cases\TC010_受注取込_0件\input\   ← 把投入文件放这里
expected\TC010_受注取込_0件\      ← 期待值放这里
```

四种模板（`--template`）：

| 名字 | 用途 |
|---|---|
| `normal` | 正常系（既定） |
| `error` | 异常系：不正数据 / 坏文件 |
| `env_missing` | 环境不备：文件夹不存在等 |
| `manual` | 手动实施用例（见 6.4） |

`copy` 会把**投入文件、期待值、fixtures 一起复制**，并把 YAML 里所有旧 ID 的引用
（`expected/<旧ID>/...`）改写成新 ID。所以「同一个 batch 换一组输入」的用例，复制完只要
换掉 `input\` 里的文件、改一下期待值就行。

> **新建的用例是 `enabled: false`**。这样作到一半的用例不会把全套的 `validate` / `run`
> 拖红。准备好了再把它改成 `true`。
>
> ID 会做检查：重复、含路径分隔符、超过 31 字（Excel sheet 名上限）都会当场报错，
> 不会生成半成品。既有文件绝不覆盖。

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

### 按机能分文件夹（文件夹名自动成为 tag）

```
cases/
├─ 受注/                          ← 这个文件夹下的用例自动带 tag「受注」
│   ├─ TC_ORDER01.yaml
│   ├─ TC_ORDER01/input/...
│   └─ TC_ORDER02.yaml
├─ 請求/                          ← 自动带 tag「請求」
│   ├─ TC_INV01.yaml
│   └─ TC_INV01/input/...
└─ 共通/
    └─ TC_COMMON01.yaml
```

```cmd
run_test.bat --tag 受注           :: 只跑受注机能
run_test.bat --tag 受注 --tag 請求 :: 两个机能
```

嵌套也可以，每层都是一个 tag：`cases/受注/取込/TC001.yaml` → `[受注, 取込]`。
YAML 里 `tags:` 声明的标签会追加在后面，两者都能用来筛选。

### 不同机能调不同的 .exe

`settings.local.yaml` 里给每个 .exe 定义一个名字：

```yaml
batch:                                    # 既定
  exe_path: "C:/app/order/OrderBatch.exe"
  console_encoding: "cp932"
  timeout_sec: 600

batches:                                  # 只写和既定的差异
  invoice:
    exe_path: "C:/app/invoice/InvoiceBatch.exe"
    working_dir: "C:/app/invoice"
    log_dir: inv_log_dir
  nightly:
    exe_path: "C:/app/batch/NightlyBatch.exe"
    timeout_sec: 3600
```

用例里选：

```yaml
# cases/請求/TC_INV01.yaml
execute:
  batch: invoice                          # ← 用哪个 .exe
  args: ["--mode", "monthly"]
collect:
  folder_evidence: [inv_input_dir, inv_output_dir]   # 该机能的文件夹
```

一次执行可以跨机能，各用例用各自的 .exe，最后汇总到同一份 Excel。

打错 ID 或标签会**报配置错误并退出码 2**，不会静默跑 0 件。

**筛选执行也照常出 Excel**，而且サマリ里会记录筛选条件和未执行的件数：

```
実行対象     | タグ指定: 異常系
実行ケース数 | 3 件（全 6 件中。3 件は今回実行していません）
総合判定     | OK
```

这点对エビデンス很重要——不写的话，别人看到「3 件全 OK」会误以为全部测试都通过了。

想让文件名也带上筛选条件，改 `settings.local.yaml`：

```yaml
excel:
  file_name_format: "TestEvidence_{run_id}_{filter}.xlsx"
  # → TestEvidence_20260805_084441_468_tag-異常系.xlsx
```

```cmd
:: --- 接真实环境前的确认 ---
python -m autotest dbcheck            :: SQL Server 能不能连上 + 实际连到的是不是那个库
python -m autotest validate           :: 配置和用例的完整性（有问题退出码 1）
python -m autotest run --dry-run      :: 不碰 DB、不启动 exe，只验流程

:: --- 正式执行 ---
run_test.bat                          :: 全用例
run_test.bat --case TC001_normal      :: 指定用例（可多次指定）
run_test.bat --tag 正常系              :: 按标签
run_test.bat --dry-run                :: 不启动 exe、不动文件夹，只验流程
run_demo.bat                          :: 沙箱空跑

:: 配置文件由工具自动选：有 config\settings.local.yaml 就用它，否则用 settings.yaml
:: run_test.bat 不再固定 --config，所以 local 配置一定会生效

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

**退出码**：全 OK = `0`，有 NG = `1`，设定错误 = `2`，**有要確認 = `3`**
→ 可直接接 タスクスケジューラ / Jenkins / GitLab CI。
`run_test.bat` 从调度器启动时不会 `pause`，手动双击时才停留等按键。

某个用例抛异常不会中断整轮，会记为该用例的 `ERROR` 并继续跑后续用例。

---

### 跑不了自动化的用例：跳过自动执行，但照样取证迹

有些用例自动跑不了——执行中途要人工介入、要卡时机、要在另一台机器上操作、
要中途停下来看状态。这类用例在 YAML 里写 `mode: manual`：

```yaml
id: TC008_締め処理_手動
name: "手動：締め処理を止めながら確認"
mode: manual          # ← 这一行
```

之后它就**从 `autotest run` 里被排除**，但证迹照取，分两段：

```cmd
:: ① 执行前：清文件夹 → setup SQL → 投入数据 → 拍 DB 快照 → 拍文件夹一览
python -m autotest manual --case TC008_締め処理_手動 --phase before

:: ② 你自己手动跑 batch（爱怎么操作怎么操作，隔几分钟、换个终端都行）

:: ③ 执行后：切日志增量 → 回收成果文件 → 拍 DB 快照 → 拍文件夹一览 → 判定
python -m autotest manual --case TC008_締め処理_手動 --phase after
```

画面上就是「手動：実行前を採取」和「手動：実行後を採取」两个按钮。

中间状态存在 `output\manual_<用例ID>_<日时>\session.json`（日志字节偏移、基准日、
执行前快照的 CSV）。所以两步之间**隔多久都行，关掉终端也行**。

**判定一定是「要確認」，绝不会自动变成 OK**：

| 项目 | 行为 |
|---|---|
| 自动比对（DB / 文件 / 日志） | 照常做，但结果一致也只给「要確認」——人手动跑的，一致只是参考 |
| 不一致 | 照常 **NG**（不等人确认，直接当问题） |
| `exit_code` | 一律 SKIP。手动起动取不到退出码，**也不接受你自己填**（填了就等于用没法验证的值发合格证） |
| 退出码 | `3`（要確認） |

Excel 里那一 sheet 的判定栏是黄色的，人看完证迹在里面写结论。

> `mode: manual` 的用例不能用 `setup.replace_files` 和 `setup.db_lock`。
> 这两个功能靠自动执行的 `finally` 保证「一定还原 / 一定释放」，而 before 和 after
> 是两个进程，中间没有 finally 可依赖。要用就手动配置、手动还原（preflight 会拦下来提醒）。

**自动跑的那轮 Excel 里会写明手动用例没跑**，避免「自动分全 OK」被读成「全部通过」：

```
実行ケース数     | 6
手動実施ケース   | 1 件 未採取（autotest manual で採取してください）: TC008_締め処理_手動
```

### 把自动分和手动分合成一册提出用证迹

```cmd
python -m autotest report ^
    --run 20260807_154253_817 ^
    --run manual_TC008_締め処理_手動_20260807_154307_566 ^
    --out output\提出用.xlsx
```

`--run` 写 `output\` 下的文件夹名，可以写任意多个。合出来的一册包含全部用例的 sheet、
全部证迹图片，サマリ里记录了结合来源。已经采集过的手动用例会自动从「未採取」里消失。

> 合并是**从头重建**一个新工作簿，不是往既有的 xlsx 里追加。
> openpyxl 2.5（Anaconda 5.2 同梱）打开带图片的 xlsx 再保存会**丢掉全部图片**，
> 追加方式等于把证迹悄悄毁掉。所以每次 run 都会在 `output\<run_id>\results\` 下
> 存一份可重建的判定结果，合并时读它。
>
> 附带好处：某一轮跑到一半崩了，已完成的用例结果也还在，可以直接 `report` 出册。

**退出码**和 run 一致：`0`=全 OK、`1`=有 NG、`3`=有要確認。

---

## 7. 目录结构

```
AUTO_TEST_BATCH\
├─ setup_windows.bat      初次安装
├─ run_gui.bat            操作画面（一览 / 执行 / 建用例 / 手动采集）
├─ run_test.bat           正式执行
├─ run_demo.bat           沙箱空跑
├─ templates\             新建用例的模板（normal / error / env_missing / manual）
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
| `cli.py` | 命令行入口（`run` / `validate` / `list` / `new` / `copy` / `manual` / `report` / `gui`） |
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
| `gui.py` | 操作画面（tkinter）。执行是 subprocess 调 CLI，不另起编排路径 |
| `scaffold.py` | 从模板建用例 / 复制既有用例 |
| `manual.py` | 手动实施用例的 before/after 中间状态（session.json） |
| `result_store.py` | 判定结果的序列化。`report` 靠它重建合并的 Excel |

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
| **个人信息** | `snapshot.tables[].mask` 列会替换成 `***MASKED***` 后再写进 Excel，证据簿可直接共享。**但遮蔽只发生在写 Excel 时**：`output\<run_id>\results\*.json`、`output\manual_*\before\*.csv` 和 `artifacts\` 里都是原值（比对必须用原值，否则值不同也会看起来相同）。**能对外共享的只有 .xlsx**，run 文件夹本身按内部资料管理 |

---

## 9. 后续可扩展方向

- **套用公司 Excel 模板**：现在是从零生成。若有固定的エビデンス様式，改 `excel.py` 的
  `build_workbook` 为「打开模板 → 往指定位置填」即可，其余模块不用动。
- **跨表判定**：目前 DB 判定是单表对单 CSV。需要跨表校验时，在 `snapshot.tables` 里
  直接写 JOIN 的 SQL，把结果当成一张虚拟表比对。
- **性能回归**：`ExecutionInfo.elapsed_sec` 已记录，加个阈值断言即可。

---

## 10. Excel → Layout TXT 生成工具

用于从帐票 Layout 定义 Excel 自动制作 TXT 取入数据。Windows 上双击：

```text
run_layout_txt.bat
```

界面选择 Excel、sheet 和输出目录即可。默认列与照片中的定义一致：

GUI操作有两条路径：

1. 点击「Excel定义を読込」，从下拉框选择一个 `FORM_ID`。画面会显示该FORM的
   Excel行、LAYOUT_ID、ELEMENT_ID、ITEM_NAME、DATA_TYPE、IME、MAX和自动生成的OCR值。
2. 双击表格中的 `出力`、`ELEMENT_ID`、`OCR生成値`、`属性`、`座標`可以修改；
   点击「表示中FORM_IDを出力」只生成当前FORM和画面上启用的项目。
3. 不需要预览修改时，点击「全FORM_IDを直接出力」，按Excel定义一次生成全部FORM。

「生成・ファイル設定」以及Excel路径、列指定、表格显示列会自动保存。也可以点击
「設定を保存」立即写入配置文件；下次启动源码版或EXE版时自动恢复。Windows默认保存位置：

```text
%APPDATA%\AUTO_TEST_BATCH\layout_txt_gui.json
```

如需固定到其他位置，可设置环境变量 `AUTOTEST_LAYOUT_GUI_CONFIG`。配置使用UTF-8 JSON，
并以临时文件替换方式写入，避免程序中断留下半个配置文件。

TXT文件名可以直接写固定名称，也可以使用 `{form_id}`、`{pattern}`、`{seq:02d}`、
`{source}` 模板。选择单个FORM时可以写成 `TEST_1001`；批量生成时建议保留
`{form_id}`，避免不同FORM重名。

| 用途 | 默认列 | 示例值 |
|---|---:|---|
| FormID | B | `1001` |
| LayoutID（布局识别） | C | `00001` |
| ITEM_NAME | H | `傷病名1` / `生年月日` |
| 数据类型 | I | `文字列` / `日付` / `カレンダー` / `チェックボックス` |
| IME / 输入限制 | J | `ひらがな` / `数値のみ` / `小数点数値` / `半角カタカナ` / `半角英数` / `全タイプ` |
| 最大位数 | K | `100` / `NULL` |
| FieldID | L | `1001` / `1002`（`ELEMENT_ID`） |

还可以指定4个非必需列：`入力属性`、`入力規則`、`補足`、`出力例`。默认值为 `auto`，
存在同名表头时自动读取并追加显示在表格末尾；不存在时不会报错。也可以填写Excel列字母、
准确表头名，或填写 `none` 禁用。这4列只用于画面确认，不会改变TXT取入格式。

点击「表示列...」可以分别显示或隐藏表格中的任何一列，选择会保存到上述配置文件。
每行的「出力」列仍用于决定当前FORM输出哪些项目，双击可单独切换；批量的
「全項目ON/OFF」按钮已删除。

FormID、LayoutID 和 FieldID 会根据表头自动寻找：FormID 使用 B 列 `FORM_ID`，
LayoutID 使用 C 列 `LAYOUT_ID`，FieldID 优先使用 L 列 `ELEMENT_ID`。
**不会拿 `ITEM_NAME` 代替任何 ID**；
界面里也可以填 Excel 列字母或准确表头名覆盖自动判断。数字 ID 若使用 `00000` 这种 Excel
显示格式，会保留为 `00001`，不会变成 `1`。

可生成三种数据：

- `通常値`：生成不超过 K 列限制的代表性合法值。
- `最大桁数ちょうど`：生成长度等于 K 的边界值。
- `最大桁数 + 1`：生成超过 K 一位的异常测试值。

具体值生成规则：

- 普通 FormID：生成全账票正常数据，所有属性Flag固定为 `0`（正常识别），日期使用
  `5/8/6/1`（和暦）。
- `FORM_ID=4001`：默认生成24个独立的正常/异常Pattern，每个Pattern分别输出
  `TXT + 同名TIF`。属性含义为 `0`（正常识别）、`1`（个数不对）、
  `2`（识别不可）、`1,2`（1和2混在）。`1,2`使用半角逗号；由于整个值有双引号，
  因此仍是一个属性字段。
- `日付` / `日付 (From)` / `日付 (To)` 始终只生成一个日期，4001中按顺序循环3种形式：
  `5/8/6/1`（和暦）、`/2026/6/1`（西暦）、`5/2026/6/1`（元号+4位西暦）。
  即使GUI选择「複数」，这3种数据类型也不会出现 `|`。
- `カレンダー`：Excel中的每一行是一个独立项目。例如Excel定义46行カレンダー时，
  TXT生成46个项目数据块，不会把46个日期塞进一个项目。每个项目都允许使用单日期，
  或使用 `5/8/6/1|5/8/6/2` 这种 `|` 分隔的多个日期；4001中循环上述3种单日期
  加这一种复数日期，共4种形式。
- `選択肢`：按项目依次循环生成 `1`、`2`、`1.2`。
- `チェックボックス`：只能生成 `0` 或 `1`，按项目交替生成。
- `文字列`：不添加分隔符，优先使用该行 `ITEM_NAME`，并按 K 列最大长度处理。
- `氏名`：根据 `ELEMENT_IME_NAME` 生成；`数値のみ` 为纯数字，
  `小数点数値` 使用 `1.0` 形式，其他IME也遵守各自输入限制。
- `座標情報`：每个FORM分别从 `0,0,0,1` 开始，只对最后一位连续编号；下一个FORM
  重新从1开始。即使错误Pattern增加或复制项目，也会继续分配末尾编号，保证FORM内不重复。

I 列为 `文字列` 时，`通常値` 优先使用该行的 `ITEM_NAME` 作为 OCR 测试值，并按 K 列
最大长度截断；`最大桁数` 和 `最大桁数 + 1` 模式则以 `ITEM_NAME` 为种子补足目标长度。
日期、checkbox、选择项等非文字类型仍生成对应格式的数据。

普通 FormID 默认输出一组同名文件，例如 `1001.txt + 1001.tif`。全网罗 FormID 4001
默认输出24组，例如 `4001_01_normal.txt + 4001_01_normal.tif`、
`4001_03_count_missing.txt + 4001_03_count_missing.tif`。TXT使用 `cp932 + CRLF`；
TIF为A4相当、200 DPI，显示该Pattern实际使用的FormID、ELEMENT_ID、ITEM_NAME、OCR值和属性，
项目过多时会生成多页TIF。
勾选「TARを生成」会把本次的TXT/TIF一起放入标准 `.tar`；勾选「TARだけ残す」时，
输出目录只留下TAR，不留下散装TXT/TIF。TAR文件名支持 `{source}` 和 `{form_id}`；
选择全部FORM时 `{form_id}` 会变成 `all`。
一个账票的全部数据只占一行，
每个值都用双引号包裹并用逗号分隔，顺序为：

```text
"FormID","対象有無","FieldID","OCR文字識別結果","属性フラグ","座標情報",...
```

例如，普通账票 `1001.txt`：

```text
"1001","1","1001","傷病名1","0","0,0,0,1","1002","1","0","0,0,0,2"
```

全网罗账票的正常Pattern `4001_01_normal.txt`：

```text
"4001","1","2001","5/8/6/1","0","0,0,0,1","2002","/2026/6/1","0","0,0,0,2","2003","5/2026/6/1","0","0,0,0,3","2004","5/8/6/1|5/8/6/2","0","0,0,0,4"
```

4001的24个Pattern如下：

| 分类 | Pattern | 实际生成内容 |
|---|---|---|
| 正常 | `normal` | 所有定义Field、属性0 |
| 个数不对 | `count_zero` / `count_missing` / `count_extra` / `count_duplicate_all` | 0件、少1件、多1件、全体重复；属性1 |
| ELEMENT_ID | `element_id_unknown` / `element_id_duplicate` / `element_id_empty` / `element_order_reverse` | 定义外、重复、空、顺序反转 |
| 属性 | `unrecognizable` / `attribute_mixed_1_2` / `attribute_1_without_count_error` / `attribute_invalid` | 属性2、`1,2`混在、件数正常但属性1、未定义属性9 |
| OCR值 | `ocr_empty` / `ocr_over_max` / `ocr_invalid_date` / `ocr_invalid_selection` / `ocr_invalid_ime` | 空值、最大位数超限、错误日期、非数字选择值/错误分隔、IME限制外字符 |
| 坐标 | `coordinates_empty` / `coordinates_invalid` | 空坐标、非数字/负值坐标 |
| 账票头 | `target_absent` / `target_empty` / `form_id_mismatch` / `form_id_empty` | 对象外、对象有无为空、FormID不一致、FormID为空 |

其中“个数不对”不只是把属性改成1：TXT中的Field数据块数量也会真的减少、增加或重复。
`ELEMENT_ID`不匹配则使用独立Pattern，便于明确判断系统报错原因。

文件名可通过模板指定。支持 `{form_id}`、`{output_form_id}`、`{pattern}`、`{seq:02d}`、
`{source}`。例如 `CASE_{form_id}_{seq:02d}_{pattern}`。模板没有写 `{pattern}` 或 `{seq}`
但需要输出多个Pattern时，工具会自动追加序号和Pattern名，防止覆盖。TXT和TIF始终使用相同basename。

默认采用最适合直接取入的“一账票一行”格式；可先选 `ラベル付き（確認用）` 核对内容。
如果实际接口的属性 flag、坐标有固定格式，
命令行可以覆盖这些默认值：

```cmd
set PYTHONPATH=src
python -m autotest.layout_txt layout.xlsx ^
    --sheet 帳票対象整理 ^
    --out-dir output\layout_txt ^
    --form-column B --layout-column C --field-column L ^
    --date-mode coverage --coverage-form-id 4001 ^
    --error-patterns all ^
    --filename-template "CASE_{form_id}_{seq:02d}_{pattern}" ^
    --profile normal --encoding cp932 ^
    --attribute-flag 0 --coordinates auto
```

只输出FORM_ID 1001、指定TXT名并且只保留TAR：

```cmd
set PYTHONPATH=src
python -m autotest.layout_txt layout.xlsx ^
    --sheet 帳票対象整理 ^
    --out-dir output\layout_txt ^
    --form-id 1001 ^
    --error-patterns none ^
    --filename-template "TEST_1001" ^
    --tar-only --tar-name "FORM_1001_DATA"
```

多个FORM可以重复指定或用逗号连接：`--form-id 1001 --form-id 1003`、
`--form-id 1001,1003`。省略 `--form-id` 就是直接输出全部。

已有同名 TXT/TIF/TAR 时默认报错并保持原文件；明确传 `--overwrite` 才覆盖。
不需要TIF时使用 `--no-tif`；只生成主要8种异常时使用 `--error-patterns core`；
不生成异常Pattern时使用 `--error-patterns none`。运行 `python -m autotest.layout_txt --help`
可查看单文件输出、TSV、UTF-8、自定义 I/J/K 列等选项。

### 构建Windows EXE

在目标Windows机器的 **Anaconda Prompt** 中执行：

```cmd
build_layout_exe.bat --install
```

首次执行会安装固定版本 `PyInstaller 4.10`，然后构建并自动做依赖冒烟检查。完成后的程序位于：

```text
dist\LayoutTxtGenerator\LayoutTxtGenerator.exe
```

这是稳定性优先的 `onedir` 版本。交付时请复制整个
`dist\LayoutTxtGenerator` 文件夹，不能只复制其中的EXE。目标电脑不需要另外安装Python。
以后构建环境不变时直接执行 `build_layout_exe.bat` 即可。

如果要明确使用Anaconda 5.2的Python：

```cmd
set LAYOUT_BUILD_PYTHON=C:\ProgramData\Anaconda3\python.exe
build_layout_exe.bat --install
```

PyInstaller不是交叉编译器，因此Windows EXE必须在Windows上构建；macOS/Linux只能验证源码和
spec，不能产出可运行的Windows EXE。构建环境支持Python 3.6～3.10，本项目基准为
Anaconda 5.2 / Python 3.6.5。

推送上述相关文件到 `main` 后，GitHub Actions的 `Build Layout Generator EXE` 也会自动
在Windows上执行同样的测试、构建和冒烟检查。成功后可以从该次Actions运行的
`Artifacts` 下载 `LayoutTxtGenerator-windows-x64`，解压后双击
`LayoutTxtGenerator.exe`，有效期默认30天；也可以在Actions页面手动执行该workflow。
