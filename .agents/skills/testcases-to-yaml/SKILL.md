---
name: testcases-to-yaml
description: 把已经写好的 autotest 测试用例 md（テストケース md）用 autotest gen 确定性地转换成用例 YAML・投入文件・初期数据 SQL・期待值，修正 gen 报出的格式错误并做 preflight。Use when a testcase markdown already exists and needs to become cases/*.yaml, or when gen reports 形式エラー.
---

# 测试用例 md → 用例 YAML

对象：用户指定的测试用例 md。YAML 不要手写——转换由 `autotest gen` 完成（不调用 LLM）。

1. `PYTHONPATH=src python -m autotest gen --md <md> --dry-run --json`（只读这 1 行 JSON）
2. `errors` 不为空时：只按 `errors[].line` 修改 md 的对应行（规则见 `docs/ai/testcase-format.md`，只读需要的节），
   回到 1。不要整份重写 md，也不要读生成的 YAML。
3. 0 错误后去掉 `--dry-run` 执行（仍加 `--json`）；md 是改过的旧用例就加 `--overwrite`（手改过的 YAML 会显示「対象外」，不会被覆盖）。
   退出码 1 / `blocked` 非空 = 有用例没反映 md：只对 `next` 里的用例继续，把 `blocked` 报告给用户。
4. `PYTHONPATH=src python -m autotest validate --include-disabled --case <ID>...`，把 NG 和 gen 的「警告 / 残作業」整理给用户。
5. 不要把 `enabled` 改成 `true`（人レビュー后再改）。Windows：cmd `set PYTHONPATH=src` / PowerShell `$env:PYTHONPATH="src"`。
