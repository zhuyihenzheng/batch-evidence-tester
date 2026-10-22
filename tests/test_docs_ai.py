# -*- coding: utf-8 -*-
"""AI 代理向けドキュメント（AGENTS.md / docs/ai / skills）が実装とずれないことの確認。

AI はドキュメントに書かれた項目しか使わない。項目を追加したのにドキュメントへ
書き忘れると「機能はあるのに使われない」、逆に消した項目が残ると「書いたのに効かない」
定義が生成される。どちらもここで止める。

  python -m unittest tests.test_docs_ai -v
"""

import re
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import config  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SKILL_DIRS = (ROOT / ".claude" / "skills", ROOT / ".agents" / "skills")


def read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


class TestCaseSchemaDoc(unittest.TestCase):
    KEY_SETS = ("CASE_TOP_KEYS", "SETUP_KEYS", "SETUP_BATCH_KEYS", "CLEAN_DIR_KEYS", "SQL_ITEM_KEYS",
                "INPUT_FILE_KEYS", "CORRUPT_KEYS", "REPLACE_FILE_KEYS", "SNAPSHOT_KEYS", "SNAPSHOT_TABLE_KEYS",
                "COLLECT_KEYS", "COLLECT_FILE_KEYS", "EXECUTE_KEYS", "ASSERT_KEYS", "DB_ASSERT_KEYS",
                "FILE_ASSERT_KEYS", "FILE_ACTUAL_KEYS", "FILE_EXISTS_KEYS", "LOG_ASSERT_KEYS",
                "EXIT_CODE_ASSERT_KEYS", "JOB_STATUS_ASSERT_KEYS", "TEARDOWN_KEYS", "BATCH_PROFILE_KEYS",
                "DATABASE_SQLSERVER_KEYS", "DATABASE_POSTGRESQL_KEYS")

    def test_every_allowed_key_is_documented(self):
        doc = read("docs/ai/case-schema.md")
        missing = []
        for name in self.KEY_SETS:
            for key in sorted(getattr(config, name)):
                if not re.search(r"(?<![A-Za-z_])%s(?![A-Za-z_])\s*:" % re.escape(key), doc):
                    missing.append("%s.%s" % (name, key))
        self.assertEqual(missing, [], "docs/ai/case-schema.md に書かれていない許可キー")

    def test_removed_key_is_not_documented_as_usable(self):
        doc = read("docs/ai/case-schema.md")
        self.assertNotIn("decimal_places", doc)
        # expected_exit_code は setup.batches（前置 batch）でだけ有効。execute には書けない
        lines = [line for line in doc.splitlines() if "expected_exit_code" in line]
        self.assertTrue(lines)
        self.assertTrue(all("batch:" in line for line in lines), lines)


class TestAgentEntryFiles(unittest.TestCase):
    def test_agents_md_stays_small_and_routes_to_existing_docs(self):
        text = read("AGENTS.md")
        # 常駐コンテキストに載るファイル。肥大化はそのまま毎回のトークン消費になる
        self.assertLess(len(text.encode("utf-8")), 8 * 1024)
        for rel in sorted(set(re.findall(r"`(docs/ai/[\w.-]+\.md)`", text))):
            self.assertTrue((ROOT / rel).is_file(), rel)

    def test_claude_md_imports_agents_md(self):
        self.assertIn("@AGENTS.md", read("CLAUDE.md"))

    def test_copilot_instructions_point_to_agents_md(self):
        text = read(".github/copilot-instructions.md")
        self.assertIn("AGENTS.md", text)
        self.assertLess(len(text), 4000, "Copilot のコードレビューは先頭 4000 文字しか読まない")

    def test_skills_are_valid_and_identical_across_tools(self):
        names = sorted(p.name for p in SKILL_DIRS[0].iterdir() if p.is_dir())
        self.assertEqual(names, ["spec-to-testcases", "testcases-to-yaml"])
        for name in names:
            texts = [(base / name / "SKILL.md").read_text(encoding="utf-8") for base in SKILL_DIRS]
            self.assertEqual(texts[0], texts[1], "%s: .claude と .agents の内容が違う" % name)
            front = re.match(r"---\nname: (?P<name>[\w-]+)\ndescription: (?P<desc>.+?)\n---\n", texts[0], re.S)
            self.assertIsNotNone(front, name)
            self.assertEqual(front.group("name"), name, "name はフォルダ名と一致させる")
            self.assertLess(len(front.group("desc")), 1024)
            for rel in re.findall(r"`(docs/ai/[\w.-]+\.md)`", texts[0]):
                self.assertTrue((ROOT / rel).is_file(), rel)


if __name__ == "__main__":
    unittest.main()
