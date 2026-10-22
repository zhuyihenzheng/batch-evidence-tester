# AGENTS.md — AUTO_TEST_BATCH

给 AI 编码代理（Codex / GitHub Copilot / Claude Code 等）的项目说明。人读的完整手册是 `README.md`
（中文、约 70KB）——**不要整份读**，按下面的路由只读需要的文档。

## 这是什么

- **autotest**（主体）：Windows 上无界面 batch 的自动测试 + 证迹（エビデンス）Excel。
  被测 batch 有两种：C# `.exe`（`batch.type: exe`，默认）和 Java Spring Boot 可执行 jar（`batch.type: java`）。
  DB 是 SQL Server 或 PostgreSQL（`database.type`；PostgreSQL 用内置客户端，不装驱动）；macOS 上用 `--offline`
  （`fixtures/` 代替 DB）开发。入口 `python -m autotest`。
- **layout_txt**：Excel 布局定义 → OCR 取込测试用 TXT/TIF/TAR，独立 GUI / EXE。与 autotest 的用例体系无关。

## 任务路由（先读对应文档）

| 任务 | 读 | 不要读 |
|---|---|---|
| 从式样书做测试用例（含测试数据、期待值） | `docs/ai/testcase-workflow.md`（它会指定再读哪两份） | README.md、`src/`、`templates/` |
| 测试用例 md → 用例 YAML | `docs/ai/testcase-workflow.md` §3 —— 跑 `autotest gen`，不要手写 YAML | 同上 |
| 直接写 / 改用例 YAML、settings | `docs/ai/case-schema.md` | |
| 改 `src/` `tests/` `demo/` `tools/` 的代码 | `docs/ai/development.md` | |
| layout_txt / EXE 打包 / CI | `docs/ai/development.md` 的「layout_txt」「EXE 打包与 CI」节 | |

## 硬性约束（违反时不会马上报错，所以写在这里）

- **Python 3.6 兼容**（目标环境 Anaconda 5.2：Python 3.6.5 / PyYAML 3.12 / openpyxl 2.5）。不用 dataclass、
  海象运算符、`list[...]` 标注、`subprocess(capture_output=/text=)`、`Path.is_relative_to`、`str.removeprefix`、
  `yaml.dump(sort_keys=)`。改 `src/` 或 `demo/` 后必须跑 `python tools/check_py36_compat.py src/autotest demo`。
- **不 `pip install -e .`**（封闭网络）。一律 `PYTHONPATH=src`。依赖只用 `requirements.txt`，不新增第三方库。
- **退出码是对外契约**：`0` 全 OK / `1` 有 NG / `2` 设定错误 / `3` 有「要確認」。
- 语言：代码注释、YAML 注释、CLI 输出、Excel、测试用例 md 的内容用**日文**；README 和 `docs/ai/` 用中文。
  注释写「为什么」，不写「做了什么」。
- 测试一律标准库 unittest 风格（CI 另用 pytest 跑同一批文件）。

## 第一原则：偽 OK 是最重的缺陷

本工具是判定别人对不对的裁判。「本来 NG 却报 OK」比崩溃严重得多。**歧义一律往 NG / REVIEW 倒，绝不往 OK 倒**。

- 不整形、不截断比较用的值（前后空白也是值）；表示用的打切不进判定路径。
- 执行 0 件、全 SKIP 都不算成功；筛选执行和未执行的手动用例要在证迹里留痕。
- 用例 YAML（含下位 mapping）的**未知键、重复键是硬错误**；什么都不检查的断言（空的 `assert.log` 等）是错误。
- `manual: true` / `mode: manual` 一致也只到 REVIEW；自动 NG 不能被人工覆盖。
- Spring Boot 的退出码不可信（作业 FAILED 也可能是 0）→ 用 `assert.job_status`。
- 破坏性操作前先跑 `preflight_case`；清空只经过 `fsops.assert_safe_to_clear`。细节见 `docs/ai/development.md`。

## 用例与生成物（AI 最容易犯的）

- 新用例：写测试用例 md → `autotest gen`。生成的 YAML 第 1 行是 `# generated-by:` 标记，
  **要改就改 md 再 `gen --overwrite`**，直接改 YAML 会让它脱离 gen 的管理。
- `enabled: false` = 未レビュー。**不要擅自改成 `true`**（用户明确要求时除外）。`run` 不执行它，
  检查草稿用 `validate --include-disabled`。
- 期待值从式样书推导；推导不出就写进人工确认（「確認」/ `review:`）并在回复里说明，不要编造。
- 日志文言、退出码只用式样书里的原文 / 定义值。

## 常用命令

```bash
python -m unittest discover -s tests                                         # 全部测试
python tools/check_py36_compat.py src/autotest demo                          # 3.6 兼容检查
PYTHONPATH=src python -m autotest run --config config/settings.demo.yaml --offline   # 沙箱全流程（TC003 故意 NG → 退出码 1）
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --dry-run --json  # 测试用例 md 的格式检查（1 行 JSON）
PYTHONPATH=src python -m autotest gen --md docs/testcases/x.md --check      # md 与生成物是否一致（不一致 = 1）
PYTHONPATH=src python -m autotest validate --include-disabled                # 含未启用草稿的 preflight
PYTHONPATH=src python -m autotest list --include-disabled
```

Windows：cmd 用 `set PYTHONPATH=src`，PowerShell 用 `$env:PYTHONPATH="src"`，再执行 `python -m autotest ...`
（`.bat` 脚本已内置）。配置文件自动选取：`config/settings.local.yaml`（各人实环境，gitignore）优先于 `config/settings.yaml`。

## 节省 token

- 按上面的路由读文档；看大文件先 `grep -n '^#' <file>` 再读需要的节。
- YAML 用 `gen` 生成，不要手写。转换本身不调用 LLM，耗 token 的是读输出 → 用 `gen --json`，只读那 1 行；
  **不要打开生成的 YAML 去确认**。交还前 `gen --check`（md 与生成物一致才退出码 0）。
- 修 `gen` 的错误时按 `L<行号>` 局部改 md，不要整份重写。
- 回复用户时只报结论、数字和需要人决定的事项，不复述文档内容。
