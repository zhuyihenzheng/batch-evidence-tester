# Copilot instructions — AUTO_TEST_BATCH

项目说明的正本是根目录的 `AGENTS.md`（任务路由・约束・命令）。这里只列每次都必须遵守的几条：

- **Python 3.6 兼容**（Anaconda 5.2 / PyYAML 3.12 / openpyxl 2.5）：不用 dataclass、海象运算符、`list[...]` 标注、
  `subprocess(capture_output=/text=)`、`str.removeprefix`、`yaml.dump(sort_keys=)`。改完跑
  `python tools/check_py36_compat.py src/autotest demo`。不新增第三方依赖，不用 `pip install -e .`（`PYTHONPATH=src`）。
- **偽 OK 是最重的缺陷**：歧义往 NG / REVIEW 倒，绝不往 OK 倒。不整形比较值；未知的 YAML 键是硬错误；
  什么都不检查的断言不允许。Spring Boot 用 `assert.job_status`，不只看退出码。
- 退出码契约：`0` 全 OK / `1` NG / `2` 设定错误 / `3` 要確認。
- 代码注释・CLI 输出・Excel・测试用例 md 内容用日文（写「为什么」）；README 和 `docs/ai/` 用中文。
- 测试用例：写 `docs/ai/testcase-format.md` 格式的 md → `python -m autotest gen --md <md>`。不要手写用例 YAML，
  不要把生成物的 `enabled: false` 改成 `true`。流程见 `docs/ai/testcase-workflow.md`。
- 不要通读 README.md（70KB）；按 `AGENTS.md` 的路由读 `docs/ai/` 下的对应文档。
