# 工作流：式样书 md → 测试用例 md → 用例 YAML

> 给 Claude Code / GitHub Copilot / Codex 等代理（人也照此操作）。
> 分工：**LLM 只做设计**——测试观点、测试数据、期待值都写进测试用例 md；
> **md → YAML / 数据文件 / 初期数据 SQL 由 `autotest gen` 确定性转换**（转换本身不调用 LLM，结果必然符合 schema；代理只读它输出的 1 行 JSON）。

## 0. 开始前确认（不清楚就问用户，不要猜）

- 式样书路径；输出的测试用例 md 路径（默认 `docs/testcases/<机能>.md`）；front matter 的 `出力先`（cases/ 下的组）。
- 被测 batch：C# exe（默认 batch）还是 Java Spring Boot（settings 里 `type: java` 的命名 batch → 用例写 `バッチ:`）。
- 可用的论理名：`grep -n "^  [a-z_]*:" config/settings*.yaml cases/*/settings.yaml`。md 里只能用已定义的论理名。

## 1. 读式样书（省 token）

- 先 `grep -n '^#' <式样书.md>` 看目录，只读与本次机能相关的节；表定义、消息一览、终了代码表要读全。
- 只读这三份：式样书、`docs/ai/test-design.md`、`docs/ai/testcase-format.md`。
  **不要读** README.md（70KB）、`src/`、`templates/`、已有的用例 YAML。
- 往已有的测试用例 md 追加时：`grep -n '^## ' <md>` 看已有 ID，只追加新节，不重写已有用例。

## 2. 写测试用例 md（内容用日文）

- 严格按 testcase-format.md 的格式；1 机能 1 文件；导入部放「観点一覧」表（仕様 → ケース ID）。
- 测试数据 / 期待值写成 CSV 代码块（`seed=` / `input=` / `expect_db=` / `expect_file=`）。大数据不要贴进 md。
- 日志文言、退出码、作业状态只用式样书里有原文的；推导不出的写进「確認」，并在报告里列为待确认的假设。
- Spring Boot 用例一律写「ジョブ状態」（退出码不可信）。

## 3. 转换与校验（循环到 0 错误）

```bash
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --dry-run --json   # 只检查；按 errors[].line 改 md
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --json             # 生成（全部 enabled: false）
PYTHONPATH=src python -m autotest validate --include-disabled --case <ID> ... # preflight：论理名 / batch / 期待值文件
```

- Windows：cmd 先 `set PYTHONPATH=src`，PowerShell 先 `$env:PYTHONPATH="src"`，再 `python -m autotest ...`。
- 修错只改 md 的对应行，不要整份重读重写；**不要打开生成的 YAML 去「确认」**——gen / validate 的输出就是结果。
- md 改了之后：`gen --overwrite`。手改过的 YAML / 数据文件显示「対象外」，永远不会被覆盖。
  这时 gen 退出码 = 1、`--json` 的 `blocked` 列出这些 ID —— 不要对它们跑 validate / run 当作「已按 md 验证」，
  在报告里写明，由人决定把手改内容移进 md 还是放弃手改。
- 交还前 `gen --check`：md 与生成物一致才退出码 0（差分あり / 手改 / 未生成 / md 里删掉的用例都算不一致）。
- gen 的「警告」（如：判定目录没清理 → 偽 OK 风险）要么修 md，要么在报告里说明为什么可以忽略。

## 4. 交还给人

- **不要把 `enabled` 改成 `true`**——那是人看过 md 和生成物之后的动作（用户明确要求时除外）。
- 报告（简短）：生成的用例数与 ID、観点一覧的覆盖情况、gen 的警告 / 残作業、你做的假设（尤其是期待值）。
- 人确认后：改 `enabled: true` → `run`（真环境）；没有 DB 的机器上 `run --offline` 只验配线（需要 `fixture=` 数据）。
