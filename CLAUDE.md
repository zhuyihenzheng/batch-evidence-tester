# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.
共通的项目说明在 AGENTS.md（Codex / Copilot 也读同一份），这里只放 Claude Code 专用的部分。

@AGENTS.md

## Claude Code 专用

- 项目 skill（`.claude/skills/`，Codex / Copilot 用的同内容版本在 `.agents/skills/`）：
  - `/spec-to-testcases <式样书.md>` — 式样书 → 测试用例 md（含测试数据・期待值）→ `gen` → `validate`
  - `/testcases-to-yaml <测试用例.md>` — 已有的测试用例 md → `gen` → `validate`，修格式错误
- 改代码前读 `docs/ai/development.md`（原 CLAUDE.md 的全部工程细节——架构、layout_txt、EXE/CI、
  偽 OK 防线、破坏性操作护栏、加字段的连带修改——都迁到了那里，内容未删减）。
- 常驻上下文只放 AGENTS.md 和本文件。新的细节写进 `docs/ai/`，不要再往这两个文件里堆（省 token）。
