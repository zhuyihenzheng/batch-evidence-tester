---
applyTo: "src/**/*.py,tests/**/*.py,demo/**/*.py,tools/**/*.py,*.py"
---
改代码前读 `docs/ai/development.md`（架构、偽 OK 防线、破坏性操作护栏、加字段时的连带修改）。
必须 Python 3.6 兼容（`python tools/check_py36_compat.py src/autotest demo`）；测试用标准库 unittest 风格；
注释・CLI 输出用日文，写「为什么」。
