---
applyTo: "**/testcases/**/*.md"
---
这是 `autotest gen` 的输入，格式严格按 `docs/ai/testcase-format.md`（未知标签・散文行都是错误）。
内容用日文；测试数据・期待值写成 CSV 代码块；日志文言・终了代码只用式样书原文。
改完跑 `PYTHONPATH=src python -m autotest gen --md <这个文件> --dry-run`。设计观点见 `docs/ai/test-design.md`。
