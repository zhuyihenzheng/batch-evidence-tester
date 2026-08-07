# -*- coding: utf-8 -*-
"""ケースひな形の生成・複製（autotest new / copy）の回帰テスト。

このコマンドの目的は「ケースを増やす手間を減らす」ことだが、
既存の作業中ケースを壊したり、読めない定義を黙って残したりしては本末転倒。
守るべき性質はその 2 点に集中している。

  python -m unittest discover -s tests -v
"""

import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import scaffold  # noqa: E402
from autotest.config import ConfigError, load_cases  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parent.parent


class ScaffoldCase(unittest.TestCase):
    """テンプレートを持つ最小のプロジェクトを一時フォルダに組む。"""

    TEMPLATE = (
        "# ひな形コメント（消えてはいけない）\n"
        "id: {case_id}\n"
        "name: \"正常系：（ケース名）\"\n"
        "tags: [正常系]\n"
        "enabled: false\n"
        "setup:\n"
        "  clean_dirs: [input_dir]\n"
        "assert:\n"
        "  exit_code: 0\n"
        "  db:\n"
        "    - table: T_ORDER\n"
        "      expected: \"expected/{case_id}/db_T_ORDER.csv\"\n"
        "      key: [ORDER_ID]\n"
    )

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="autotest_scaffold_"))
        self.cases = self.tmp / "cases"
        self.cases.mkdir()
        (self.tmp / "templates").mkdir()
        (self.tmp / "templates" / "normal.yaml").write_text(self.TEMPLATE, encoding="utf-8")
        (self.tmp / "templates" / "error.yaml").write_text(
            self.TEMPLATE.replace("正常系", "異常系"), encoding="utf-8")
        # 既存ケースが 1 件ある状態（複製元 / 衝突相手）
        (self.cases / "TC001_base.yaml").write_text(
            "id: TC001_base\n"
            "name: 元ケース\n"
            "tags: [正常系]\n"
            "setup:\n"
            "  input_files:\n"
            "    - {src: \"input/DATA.csv\", dest_dir: input_dir}\n"
            "assert:\n"
            "  db:\n"
            "    - table: T_ORDER\n"
            "      expected: \"expected/TC001_base/db_T_ORDER.csv\"\n"
            "      key: [ORDER_ID]\n",
            encoding="utf-8")
        (self.cases / "TC001_base" / "input").mkdir(parents=True)
        (self.cases / "TC001_base" / "input" / "DATA.csv").write_text("A,B\n1,2\n", encoding="utf-8")
        (self.tmp / "expected" / "TC001_base").mkdir(parents=True)
        (self.tmp / "expected" / "TC001_base" / "db_T_ORDER.csv").write_text(
            "ORDER_ID\n1\n", encoding="utf-8")
        (self.tmp / "fixtures" / "TC001_base").mkdir(parents=True)
        (self.tmp / "fixtures" / "TC001_base" / "before_T_ORDER.csv").write_text(
            "ORDER_ID\n1\n", encoding="utf-8")

    def tearDown(self):
        shutil.rmtree(str(self.tmp), ignore_errors=True)


class TestCreateCase(ScaffoldCase):
    def test_creates_yaml_material_and_expected_dirs(self):
        created = scaffold.create_case(self.tmp, self.cases, "TC010_new", "normal")
        self.assertTrue((self.cases / "TC010_new.yaml").is_file())
        self.assertTrue((self.cases / "TC010_new" / "input").is_dir())
        self.assertTrue((self.tmp / "expected" / "TC010_new").is_dir())
        self.assertEqual(len(created), 3)

    def test_case_id_placeholder_is_replaced_everywhere(self):
        scaffold.create_case(self.tmp, self.cases, "TC010_new", "normal")
        text = (self.cases / "TC010_new.yaml").read_text(encoding="utf-8")
        self.assertNotIn("{case_id}", text)
        self.assertIn("id: TC010_new", text)
        self.assertIn("expected/TC010_new/db_T_ORDER.csv", text)

    def test_comments_survive(self):
        """yaml.dump を通すとコメントが消える。ここが消えたら生成物の価値が落ちる。"""
        scaffold.create_case(self.tmp, self.cases, "TC010_new", "normal")
        text = (self.cases / "TC010_new.yaml").read_text(encoding="utf-8")
        self.assertIn("# ひな形コメント（消えてはいけない）", text)

    def test_generated_case_is_disabled_and_does_not_break_others(self):
        """作りかけのケースが全体の validate / run を巻き添えで落とさないこと。"""
        scaffold.create_case(self.tmp, self.cases, "TC010_new", "normal")
        ids = [c.case_id for c in load_cases(self.cases)]
        self.assertEqual(ids, ["TC001_base"], "生成直後のケースは実行対象に入らない")

    def test_duplicate_id_is_rejected_without_touching_disk(self):
        with self.assertRaises(ConfigError) as ctx:
            scaffold.create_case(self.tmp, self.cases, "TC001_base", "normal")
        self.assertIn("既に使われています", str(ctx.exception))

    def test_duplicate_id_detected_even_when_source_is_disabled(self):
        """enabled: false のケースと ID がぶつかっても証跡フォルダは衝突する。"""
        (self.cases / "TC500_off.yaml").write_text(
            "id: TC500_off\nname: 無効\nenabled: false\n", encoding="utf-8")
        with self.assertRaises(ConfigError):
            scaffold.create_case(self.tmp, self.cases, "TC500_off", "normal")

    def test_duplicate_id_detected_when_source_omits_the_id_field(self):
        """id: を書いていないケースとの衝突も検出すること。

        load_cases は id: 未指定ならファイル名を ID にする。衝突検査だけが
        それを見落とすと、生成は成功するのに次の実行で「ID 重複」で落ちる
        —— 原因の分かりにくい失敗になる。
        """
        (self.cases / "TC600_noid.yaml").write_text("name: ID なし\n", encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            scaffold.create_case(self.tmp, self.cases, "TC600_noid", "normal")
        self.assertIn("既に使われています", str(ctx.exception))
        self.assertEqual(
            (self.cases / "TC600_noid.yaml").read_text(encoding="utf-8"), "name: ID なし\n",
            "衝突時に既存ファイルを上書きしてはいけない")

    def test_invalid_id_is_rejected(self):
        for bad in ("TC/bad", "TC:bad", "..", "  "):
            with self.assertRaises(ConfigError):
                scaffold.create_case(self.tmp, self.cases, bad, "normal")

    def test_too_long_id_is_rejected(self):
        """Excel シート名の 31 文字上限を超えると見分けが付かなくなる。"""
        with self.assertRaises(ConfigError) as ctx:
            scaffold.create_case(self.tmp, self.cases, "T" * 32, "normal")
        self.assertIn("長すぎます", str(ctx.exception))

    def test_unknown_template_lists_available_ones(self):
        with self.assertRaises(ConfigError) as ctx:
            scaffold.create_case(self.tmp, self.cases, "TC010_new", "存在しない")
        message = str(ctx.exception)
        self.assertIn("normal", message)
        self.assertIn("error", message)

    def test_existing_material_dir_blocks_creation(self):
        """作成先が埋まっていたら上書きせず中止すること。"""
        (self.cases / "TC010_new" / "input").mkdir(parents=True)
        with self.assertRaises(ConfigError) as ctx:
            scaffold.create_case(self.tmp, self.cases, "TC010_new", "normal")
        self.assertIn("既に存在します", str(ctx.exception))
        self.assertFalse((self.cases / "TC010_new.yaml").exists(), "中止時に YAML を作ってはいけない")

    def test_broken_template_reports_instead_of_silently_passing(self):
        (self.tmp / "templates" / "broken.yaml").write_text(
            "id: {case_id}\nname: X\nsetupp: {}\n", encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            scaffold.create_case(self.tmp, self.cases, "TC010_new", "broken")
        self.assertIn("setupp", str(ctx.exception))


class TestCopyCase(ScaffoldCase):
    def test_copies_yaml_material_expected_and_fixtures(self):
        scaffold.copy_case(self.tmp, self.cases, "TC001_base", "TC002_copy")
        self.assertTrue((self.cases / "TC002_copy.yaml").is_file())
        self.assertEqual(
            (self.cases / "TC002_copy" / "input" / "DATA.csv").read_text(encoding="utf-8"),
            "A,B\n1,2\n", "投入ファイルごと複製されること")
        self.assertTrue((self.tmp / "expected" / "TC002_copy" / "db_T_ORDER.csv").is_file())
        self.assertTrue((self.tmp / "fixtures" / "TC002_copy" / "before_T_ORDER.csv").is_file())

    def test_id_references_inside_yaml_are_rewritten(self):
        """expected/<旧ID>/... のままだと、複製元の期待値と比較してしまう。"""
        scaffold.copy_case(self.tmp, self.cases, "TC001_base", "TC002_copy")
        text = (self.cases / "TC002_copy.yaml").read_text(encoding="utf-8")
        self.assertIn("id: TC002_copy", text)
        self.assertIn("expected/TC002_copy/db_T_ORDER.csv", text)
        self.assertNotIn("TC001_base", text)

    def test_copy_is_loadable_and_independent(self):
        scaffold.copy_case(self.tmp, self.cases, "TC001_base", "TC002_copy")
        ids = sorted(c.case_id for c in load_cases(self.cases))
        self.assertEqual(ids, ["TC001_base", "TC002_copy"])

    def test_unknown_source_lists_known_ids(self):
        with self.assertRaises(ConfigError) as ctx:
            scaffold.copy_case(self.tmp, self.cases, "TC999_nope", "TC002_copy")
        self.assertIn("TC001_base", str(ctx.exception))

    def test_duplicate_target_id_is_rejected(self):
        with self.assertRaises(ConfigError):
            scaffold.copy_case(self.tmp, self.cases, "TC001_base", "TC001_base")

    def test_source_without_id_field_can_be_copied(self):
        """id: を書かないケース（ID = ファイル名）も複製元にできること。"""
        (self.cases / "TC300_noid.yaml").write_text("name: ID なし\n", encoding="utf-8")
        scaffold.copy_case(self.tmp, self.cases, "TC300_noid", "TC301_copy")
        self.assertTrue((self.cases / "TC301_copy.yaml").is_file())

    def test_source_without_material_dir_still_gets_input_dir(self):
        (self.cases / "TC400_bare.yaml").write_text(
            "id: TC400_bare\nname: 資材なし\n", encoding="utf-8")
        scaffold.copy_case(self.tmp, self.cases, "TC400_bare", "TC401_copy")
        self.assertTrue((self.cases / "TC401_copy" / "input").is_dir())


class TestShippedTemplates(unittest.TestCase):
    """リポジトリ同梱のひな形が実際に使えること。"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="autotest_tpl_"))
        shutil.copytree(str(REPO_ROOT / "templates"), str(self.tmp / "templates"))
        (self.tmp / "cases").mkdir()

    def tearDown(self):
        shutil.rmtree(str(self.tmp), ignore_errors=True)

    def test_every_shipped_template_generates_a_valid_case(self):
        names = scaffold.list_templates(self.tmp)
        self.assertTrue(names, "templates/ にひな形が 1 つも無い")
        for i, name in enumerate(names):
            case_id = "TC%03d_%s" % (900 + i, name)
            # 生成できれば ConfigError は出ない（内部で定義の検証まで通している）
            scaffold.create_case(self.tmp, self.tmp / "cases", case_id, name)
            text = (self.tmp / "cases" / (case_id + ".yaml")).read_text(encoding="utf-8")
            self.assertNotIn("{case_id}", text, "%s に未置換のプレースホルダが残っている" % name)
            self.assertIn("enabled: false", text, "%s は enabled: false で生成すること" % name)


if __name__ == "__main__":
    unittest.main()
