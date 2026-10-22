# Spring Boot batch 测试用例生成手册（GitHub Copilot）

> 对象：要给一个新的 Spring Boot batch 做自动测试用例的人。
> 分工：**AI（Copilot / Claude Code / Codex）写测试用例 md**（观点・测试数据・期待值），
> **`autotest gen` 把 md 转换成用例 YAML**（确定性转换，不调用 LLM），**人审核后执行**。

```
式样书 md ──(Copilot)──▶ 测试用例 md ──(autotest gen)──▶ 用例 YAML（enabled: false）
                              │                                   │
                              └────────────── 人审核 ◀────────────┘ ──▶ enabled: true ──▶ autotest run
```

## 1. 准备（人来做，每个 batch 一次）

### 1.1 式样书转成 md

放到 `docs/specs/<batch名>.md`。Excel / Word 的设计书也先整理成 md，至少写清下面这些
（缺了的部分，AI 只能写成「確認」交给人判断）：

| 项目 | 例 |
|---|---|
| 作业名、step / chunk 构成 | `orderImportJob`：读取 → 校验 → 写 DB，chunk 100 件 |
| 作业参数，以及哪些是识别参数 | `targetDate`（识别）、`run.id` |
| 输入输出文件的格式 | 文件名规则、编码（UTF-8 / Shift_JIS）、换行、列定义 |
| DB 表的 DDL | 表名・列名（PostgreSQL 写实际的小写名）・类型・主键 |
| 终了代码的映射 | COMPLETED = 0 / FAILED = 1 …（没写就只靠「ジョブ状態」判定） |
| 日志文言 | 开始・结束・错误时输出的原文 |
| 异常时的规则 | 跳过上限、chunk 中途失败时的回滚范围、再实行的行为 |

### 1.2 在 settings 登记 batch 和文件夹

写在 `config\settings.local.yaml`（各人的实环境，不进 git）。各项的含义见 README 第 3 节「Java（Spring Boot）batch」。

```yaml
batches:
  order_job:
    type: java
    java: "C:/Program Files/Eclipse Adoptium/jdk-17/bin/java.exe"   # 省略 = JAVA_HOME → PATH
    jar: "C:/app/order-batch/order-batch.jar"
    jvm_args: ["-Dfile.encoding=UTF-8", "-Dspring.profiles.active=it"]
    common_args: ["--spring.batch.job.name=orderImportJob"]   # Boot 2 是 spring.batch.job.names
    log_dir: order_log_dir
paths:
  order_input_dir: "C:/app/order-batch/in"
  order_output_dir: "C:/app/order-batch/out"
  order_log_dir: "C:/app/order-batch/log"
```

DB 是 PostgreSQL 的话写 `database.type: postgresql`（README 第 3 节「PostgreSQL」）。

### 1.3 确认能找到

```cmd
python -m autotest validate
python -m autotest dbcheck
```

java、jar、文件夹都显示出实际路径，DB 显示出实际连到的服务器和库，就可以进入下一步。

### 1.4 和开发确认三件事

- **日志级别**：`org.springframework.batch` 在 INFO 以上。作业状态靠 `Job: [...] completed ... status: [...]`
  这一行判定，看不到这一行就判 NG。
- **作业参数每次都不同**：配了 `RunIdIncrementer`，或每个用例用不同的 `targetDate`。否则第二次执行会被
  `JobInstanceAlreadyCompleteException` 拒绝，作业根本不启动。
- **作业名的写法**：Boot 3 是 `--spring.batch.job.name=`，Boot 2 是 `--spring.batch.job.names=`。
  写错时 Boot 2 会把所有作业都跑一遍。

## 2. 让 Copilot 写测试用例 md 并生成

在 VS Code 的 Copilot Chat 里选 **Agent 模式**，并允许它执行终端命令（要跑 `autotest gen`）。
项目规则（`.github/copilot-instructions.md`，以及编辑 `docs/testcases/` 下的 md 时适用的规则）它会自动读，
所以提示词里只写这次的具体信息。`#file:` 是把文件附给 Copilot 的写法。

### 2.1 提示词模板（一次完成）

改掉作业名、文件名、观点即可：

```text
请用 spec-to-testcases 技能，按 docs/ai/testcase-workflow.md 的流程，为下面的 Spring Boot batch 做测试用例。

- 式样书：#file:docs/specs/order_import_job.md
- 输出：docs/testcases/order_import_job.md（front matter 的 出力先: order_import_job）
- batch：settings 的 batches.order_job（type: java），每个用例写「バッチ: `order_job`」
- 作业名：orderImportJob；作业参数 targetDate={date}（识别参数，用例之间不要重复）
- 文件夹论理名：order_input_dir / order_output_dir / order_log_dir（「クリア」里要写上）
- DB：PostgreSQL，表 t_order / t_order_detail（DDL 在式样书第 4 章，列名照 DDL 写小写）
- 观点：正常取込 / 入力チェック異常 / スキップ上限超過で FAILED / chunk 途中失敗の回滚范围 / 同一参数再実行

要求：
- 每个用例都写「ジョブ状態」（Spring Boot 的退出码不可信），终了代码按式样书 5.2 的映射
- 测试数据和期待值写成 CSV 代码块（seed= / input= / expect_db= / expect_file=）
- 日志文言、终了代码只用式样书原文；推导不出的期待值写进「確認」，不要编
- 用 gen --dry-run --json 循环到 errors 为空，再执行 gen --json、validate --include-disabled、gen --check
- 不要把 enabled 改成 true；最后简短报告：用例数、观点覆盖、gen 的警告、你做的假设
```

- 观点还没想好：删掉「观点」那一行，让它从式样书整理。
- DB 是 SQL Server：把 DB 那一行改成 SQL Server（列名照 DDL 写）。

### 2.2 式样书很长时：分两步

返工少，也更省 token。

```text
第 1 步：
读 #file:docs/specs/order_import_job.md，按 docs/ai/test-design.md 的观点清单，
只列出这个 batch 的「観点一覧」（表格：式样书节号 / 观点 / 用例 ID / 种别）。先不要写用例。
```

```text
第 2 步（确认、修改了观点一览之后）：
按上面确认过的観点一覧，用 spec-to-testcases 技能写 docs/testcases/order_import_job.md 并生成。
（这里接 2.1 模板的 batch / 作业 / DB 信息和「要求」）
```

### 2.3 式样书改了之后

```text
docs/specs/order_import_job.md 改了（看最近一次提交的 git diff）。
只修改 docs/testcases/order_import_job.md 里「仕様」节号受影响的用例，不要整份重写。
然后 gen --overwrite --json、gen --check，报告改了哪些用例、为什么。
```

### 2.4 用 Claude Code / Codex 时

- Claude Code：`/spec-to-testcases docs/specs/order_import_job.md`，后面接 2.1 模板里的信息和要求。
- Codex：2.1 的提示词直接可用（读 `AGENTS.md` 和 `.agents/skills/`）。

## 3. 审核清单（把 enabled 改成 true 之前）

- [ ] 観点一覧覆盖了式样书的各节（正常 / 异常 / 边界 / 再实行）
- [ ] 每个用例都有「ジョブ状態」；终了代码和式样书的映射一致
- [ ] 期待值是从式样书推导的，不是照抄实际运行结果（照抄等于把缺陷当成基准）
- [ ] 「確認」里是 AI 推导不出、需要人判断的事项，每条都处理过
- [ ] 测试数据最小且有意义；PostgreSQL 的列名是实际的小写名；CHAR(n) 列的期待值带填充空格
- [ ] 「クリア」包含 Spring 用的文件夹（默认只清 input / processed / error / output 四个）
- [ ] gen 的「警告」「残作業」都处理了，或说明了可以忽略的理由

没问题就把生成的 YAML（`cases/<出力先>/<ID>.yaml`）里的 `enabled: false` 改成 `true`。
之后改了 md 再 `gen --overwrite`，内容有变化的用例会自动回到 `false`，需要再审核。

## 4. 执行

```cmd
python -m autotest run --case OI_001 --case OI_002
```

也可以用 `run_gui.bat` 的操作画面勾选执行。执行结束时会打印证迹 Excel 的路径。
退出码：`0` 全部 OK / `1` 有 NG / `2` 设定错误 / `3` 有「要確認」（人工确认项・手动用例）。

## 5. 常见问题

| 现象 | 原因 | 怎么办 |
|---|---|---|
| gen「形式エラー」 | 测试用例 md 的格式问题 | 按报出的行号改 md（Copilot 会自己循环修） |
| gen 返回 1，列出「[未反映]」 | 有人手改了生成的 YAML / 数据文件，md 的修改没反映 | 把手改的内容搬回 md，删掉那个 YAML 后重新 gen；或放弃手改 |
| ジョブ状態 NG「完了ログ…が見つかりません」 | 作业没启动 / 日志级别低于 INFO / 同参数再实行被拒 | 看 1.4；提示里写了 `RunIdIncrementer` 的就是再实行被拒 |
| ジョブ状態 NG「完了行の無い実行があります」 | 作业中途停止（异常终止・被强制结束） | 看 batch 的日志和 stderr |
| ジョブ状態 要確認「複数のジョブを検出しました」 | 作业名参数没生效，跑了多个作业 | 检查 1.2 的 `common_args`；或在「ジョブ状態」里写上作业名 |
| DB 照合 NG「前後の空白だけが違う」 | CHAR(n) 列返回值带填充空格 | 快照 SQL 里 `RTRIM(列) AS 列`，或期待值补满空格 |
| PostgreSQL「列が無い」（SQLSTATE 42703） | 列名大小写和 DDL 不一致 | 测试用例 md 的 CSV 表头照 DDL 写（通常小写） |
