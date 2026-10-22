---
name: spec-to-testcases
description: 从式样书（仕様書 / 設計書）md 设计 batch 测试，写出 autotest 的测试用例 md（含测试数据・DB 初期数据・期待值），再用 autotest gen 生成用例 YAML 并校验。Use when asked to create test cases / テストケース / 测试用例 / test data / 期待値 from a spec for a C# .exe or Java Spring Boot batch in this repo.
---

# 式样书 → 测试用例 md → 用例 YAML

对象：用户指定的式样书 md（以及输出的测试用例 md 路径；没给就用 `docs/testcases/<机能>.md`）。

1. 读 `docs/ai/testcase-workflow.md` 并照它执行。除了式样书，只读 `docs/ai/test-design.md`
   和 `docs/ai/testcase-format.md`。不要读 README.md、`src/`、`templates/`、已有的用例 YAML（省 token）。
2. 式样书先 `grep -n '^#'` 看目录，只读相关节。论理名、batch 名从 settings 里 grep，不要猜。
3. 测试用例 md 用日文写；测试数据・期待值写成 CSV 代码块（`seed=` / `input=` / `expect_db=` / `expect_file=`）。
   日志文言、终了代码、ジョブ状態只用式样书里的原文；推导不出的写进「確認」。
4. `PYTHONPATH=src python -m autotest gen --md <md> --dry-run --json` 循环到 `errors` 为空（按 `errors[].line` 局部修改），
   然后去掉 `--dry-run` 生成，再 `validate --include-disabled --case <ID>...`。
5. 不要把 `enabled` 改成 `true`。最后简短报告：用例数 / 观点覆盖 / gen 的警告・残作業 / 你的假设。
