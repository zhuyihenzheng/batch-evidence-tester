# -*- coding: utf-8 -*-
"""ケース定義の下位項目の綴り間違い・何も確認しない判定を止める回帰テスト。

AI が生成した YAML で特に起きやすい「書いたのに効かない」を、実行前に止めることを守る。
あわせて validate / list --include-disabled（作成途中のケースの静的検査）を確認する。

  python -m unittest tests.test_schema_strict -v
"""

import io
import re
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import cli  # noqa: E402
from autotest.config import (ConfigError, _validate_case_schema, load_cases, load_settings,  # noqa: E402
                             load_yaml_text)

PATH = Path("cases/X.yaml")


def problems(data):
    return _validate_case_schema(data, PATH, "X")


class TestSubKeyTypos(unittest.TestCase):
    def assert_rejected(self, data, fragment):
        found = problems(data)
        self.assertTrue(any(fragment in p for p in found), found)

    def test_execute_expected_exit_code_points_to_assert(self):
        # 以前は受け付けたが何も判定していなかった（確認しているつもりになる偽 OK）
        self.assert_rejected({"execute": {"expected_exit_code": 1}}, "assert.exit_code")

    def test_snapshot_mask_typo_would_leak_pii(self):
        self.assert_rejected({"snapshot": {"tables": [{"name": "T", "sql": "SELECT 1", "masks": ["NAME"]}]}},
                             "masks")
        self.assert_rejected({"snapshot": {"tables": [{"sql": "SELECT 1"}]}}, "name がありません")
        self.assert_rejected({"snapshot": {"table": []}}, "table")

    def test_collect_preview_typo_and_string_bool(self):
        self.assert_rejected({"collect": {"files": [{"dir": "output_dir", "preveiw": True}]}}, "preveiw")
        self.assert_rejected({"collect": {"files": [{"dir": "output_dir", "preview": "true"}]}}, "true / false")

    def test_input_and_replace_file_typos(self):
        self.assert_rejected({"setup": {"input_files": [{"src": "a.csv", "dest": "input_dir"}]}}, "dest")
        self.assert_rejected({"setup": {"input_files": [{"src": "a.zip", "corrupt": {"mod": "flip"}}]}}, "mod")
        self.assert_rejected({"setup": {"replace_files": [{"src": "a", "dest_dir": "d", "nmae": "x"}]}}, "nmae")

    def test_teardown_and_sql_items(self):
        self.assert_rejected({"teardown": {"sqll": []}}, "sqll")
        self.assert_rejected({"teardown": {"sql": [{"fiel": "x.sql"}]}}, "fiel")
        self.assert_rejected({"setup": {"sql": [{"file": "a.sql", "inlne": "x"}]}}, "inlne")

    def test_remove_dirs_must_be_aliases(self):
        # 辞書を渡すと実行時に落ちていた（集合に入れられない）
        self.assert_rejected({"setup": {"remove_dirs": [{"alias": "output_dir"}]}}, "論理名の文字列")

    def test_vacuous_log_assert_is_rejected(self):
        # 文言が 1 つも無い log 判定は「必須 0 件すべて検出」で必ず OK になっていた
        self.assert_rejected({"assert": {"log": {"must_contain": []}}}, "何も確認せずに OK")
        self.assert_rejected({"assert": {"log": {"must_contain": [""]}}}, "空でない文字列")
        self.assertEqual(problems({"assert": {"log": {"manual": True}}}), [], "目視確認は許す")
        self.assertEqual(problems({"assert": {"log": {"must_not_contain": ["Exception"]}}}), [])

    def test_valid_definitions_still_pass(self):
        data = {
            "setup": {"clean_dirs": ["input_dir", {"alias": "error_dir", "exclude": ["Backup"]}],
                      "remove_dirs": ["output_dir"],
                      "sql": ["DELETE FROM T", {"file": "sql/a.sql"}, {"inline": "SELECT 1"}],
                      "input_files": [{"src": "input/a.zip", "dest_dir": "input_dir", "rename": "b.zip",
                                       "corrupt": {"mode": "truncate", "ratio": 0.5}}]},
            "snapshot": {"tables": [{"name": "T", "sql": "SELECT 1", "mask": ["NAME"]}]},
            "collect": {"files": [{"dir": "output_dir", "pattern": "*.csv", "preview": True, "recursive": False}]},
            "teardown": {"sql": [{"file": "sql/b.sql"}], "clean_dirs": ["input_dir"]},
        }
        self.assertEqual(problems(data), [])

    def test_repository_cases_and_templates_still_load(self):
        root = Path(__file__).resolve().parent.parent
        load_cases(root / "cases", include_disabled=True)
        for template in sorted((root / "templates").glob("*.yaml")):
            data = yaml.safe_load(template.read_text(encoding="utf-8").replace("{case_id}", "TC_X"))
            self.assertEqual(_validate_case_schema(data, template, "TC_X"), [], template.name)


class TestDuplicateKeys(unittest.TestCase):
    """同じキーを 2 回書いたら止めること。safe_load は後の値で黙って上書きし、消えた検査は
    スキーマ検証でも見つからない（残ったデータ自体は正しい形をしているため）。"""

    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="autotest_dupkey_"))
        self.addCleanup(shutil.rmtree, str(self.root), True)
        (self.root / "cases").mkdir()

    def test_duplicate_key_in_case_yaml_is_an_error(self):
        (self.root / "cases" / "TC001.yaml").write_text(
            "id: TC001\nname: n\nassert:\n  log: {must_contain: [処理開始]}\n  log: {must_not_contain: [ERROR]}\n",
            encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            load_cases(self.root / "cases")
        self.assertIn("'log' が 2 回", str(ctx.exception))
        self.assertIn("line 5", str(ctx.exception))

    def test_duplicate_key_in_settings_is_an_error(self):
        path = self.root / "settings.yaml"
        path.write_text("batch: {exe_path: a.exe}\npaths: {log_dir: ./log}\n"
                        "database: {server: s1, database: d}\ndatabase: {server: s2, database: d}\n", encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            load_settings(path, self.root)
        self.assertIn("'database' が 2 回", str(ctx.exception))

    def test_yaml_is_only_read_through_the_unique_key_loader(self):
        """YAML を読む箇所を増やすときに safe_load を直接使うと、重複キーの検出が抜ける。"""
        # casegen の 3 か所は利用者が書いた検査内容を読まない（生成物の読み戻し確認・enabled の値・既存 ID の収集）
        allowed = {"casegen.py": 3, "config.py": 1}
        src = Path(__file__).resolve().parent.parent / "src" / "autotest"
        found = {}
        for path in sorted(src.glob("*.py")):
            text = path.read_text(encoding="utf-8")
            count = len(re.findall(r"yaml\.(?:safe_load|load|full_load|unsafe_load)\(", text))
            if count:
                found[path.name] = count
        self.assertEqual(found, allowed, "YAML は config.load_yaml_text で読むこと")
        config_text = (src / "config.py").read_text(encoding="utf-8")
        self.assertIn("yaml.load(text, Loader=UniqueKeyLoader)", config_text)

    def test_merge_keys_and_distinct_keys_are_still_allowed(self):
        data = load_yaml_text("base: &b {a: 1, b: 2}\nx:\n  <<: *b\n  a: 3\n'1': s\n1: n\n")
        self.assertEqual(data["x"], {"a": 3, "b": 2})
        self.assertEqual((data["1"], data[1]), ("s", "n"))


class TestIncludeDisabled(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="autotest_disabled_")).resolve()
        (self.root / "config").mkdir()
        self.config = self.root / "config" / "settings.yaml"
        self.config.write_text(yaml.safe_dump({
            "env": {"sandbox": True},
            "database": {"server": "unused", "database": "unused", "auth": "windows"},
            "batch": {"exe_path": "{python}", "working_dir": "."},
            "paths": {"input_dir": "sb/in", "log_dir": "sb/log"},
        }), encoding="utf-8")
        (self.root / "cases").mkdir()
        (self.root / "cases" / "D1.yaml").write_text(
            "id: D1\nname: 作成途中\nenabled: false\nassert: {exit_code: 0}\n", encoding="utf-8")
        (self.root / "cases" / "D2.yaml").write_text(
            "id: D2\nname: 未定義 batch\nenabled: false\nexecute: {batch: nope}\nassert: {exit_code: 0}\n",
            encoding="utf-8")

    def tearDown(self):
        shutil.rmtree(str(self.root), ignore_errors=True)

    def cli(self, *args):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = cli.main(list(args) + ["--config", str(self.config)])
        return code, out.getvalue() + err.getvalue()

    def test_validate_without_flag_refuses_zero_enabled(self):
        code, out = self.cli("validate")
        self.assertEqual(code, 2)
        self.assertIn("0 件", out)

    def test_validate_include_disabled_checks_drafts_and_keeps_going(self):
        code, out = self.cli("validate", "--include-disabled")
        self.assertEqual(code, 1, "作成途中でも問題があれば NG を返す")
        self.assertIn("[OK] D1", out)
        self.assertIn("[NG] D2", out, "1 件の不備で他のケースの検査を止めない")
        self.assertIn("（無効: enabled: false）", out)

    def test_list_include_disabled_marks_state(self):
        code, out = self.cli("list", "--include-disabled")
        self.assertEqual(code, 0)
        self.assertIn("D1", out)
        self.assertIn("無効", out)

    def test_run_rejects_include_disabled(self):
        code, out = self.cli("run", "--include-disabled", "--offline")
        self.assertEqual(code, 2, "未レビューのケースを実行系で走らせない")
        self.assertIn("validate / list 専用", out)


if __name__ == "__main__":
    unittest.main()
