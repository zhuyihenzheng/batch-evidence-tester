# 式样书 → 测试用例 md → 用例 YAML 的离线样例

本文件夹是一个独立的小项目（项目根 = `demo/md_workflow/`），演示 AI 代理的完整流程：

| 文件 | 由谁产生 |
|---|---|
| `docs/specs/order_import_spec.md` | 人（式样书。C# 版 exe 与 Java/Spring Boot 版作业） |
| `docs/testcases/order_import.md` | AI（Claude Code / Copilot / Codex 按 `docs/ai/testcase-workflow.md` 写，含测试数据・期待值） |
| `cases/` `expected/` `fixtures/` | `autotest gen`（确定性转换，不要手改；改 md 后 `--overwrite`） |

在本文件夹执行（Windows 先 `set PYTHONPATH=..\..\src`）：

```sh
cd demo/md_workflow
export PYTHONPATH=../../src
python -m autotest gen --md docs/testcases/order_import.md --config config/settings.yaml --dry-run   # 格式检查
python -m autotest gen --md docs/testcases/order_import.md --config config/settings.yaml             # 生成（已生成则「変更なし」）
python -m autotest validate --include-disabled --config config/settings.yaml                        # 草稿的 preflight
```

生成物都是 `enabled: false`（未レビュー）。想在本机跑一遍时，把 `cases/order_import/*.yaml` 的
`enabled` 改成 `true`，再：

```sh
python -m autotest run --config config/settings.yaml --offline
```

预期：机器判定全部 OK；`OI_002` / `OI_004` 有人工确认项 → 要確認；`OI_005` 是手动用例，显示为未采取 →
整体「要確認」（退出码 3）。离线模式用 `fixtures/` 代替 DB，所以 DB 的比较只验证配线，不验证 batch。
试完把 `enabled` 改回 `false`（否则下次 `gen` 也能识别，只是显示「変更なし」）。

- 默认 batch = `../fake_batch.py`（C# exe 的替身）；`batches.spring` = `type: java`，用 `../fake_spring_batch.py` 代替 jar。
- `sandbox/`、`sandbox_spring/`、`output/` 是运行时生成的，已在 .gitignore 中。
- 修改了 `src/autotest/casegen.py` 的输出格式时，用 `--overwrite` 重新生成本样例（`tests.test_casegen` 会检查是否同步）。
