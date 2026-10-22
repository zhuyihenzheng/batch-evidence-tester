# -*- coding: utf-8 -*-
"""テストケース md → ケース YAML 生成（autotest gen）の回帰テスト。

守る性質は 3 つ:
  1) md に書いた検証が黙って消えないこと（未知の項目・解釈できない行・衝突はすべてエラー）
  2) 生成物が意図どおりで、そのまま読めて、手で直したものを上書きしないこと
  3) docs/ai/testcase-format.md の例と demo/md_workflow の生成物が実装とずれないこと

  python -m unittest tests.test_casegen -v
"""

import io
import json
import re
import shutil
import sys
import tempfile
import unittest
import unittest.mock
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import casegen, cli  # noqa: E402
from autotest.config import _validate_case_schema, load_cases  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parent.parent

FRONT = "---\n出力先: grp\n---\n# タイトル\n\n導入部は自由記述。\n\n"

BASIC = (
    "## TC101 正常系：基本\n"
    "- 種別: 正常系\n"
    "- 目的: 取り込まれること\n"
    "- 引数: `--mode daily --date {date}`\n"
    "- 終了コード: 0\n"
    "- ログ必須: `処理開始`, `取込件数=2`\n"
    "- ファイル:\n"
    "  - `processed_dir/ORDER_*.csv` = 1\n"
    "- 確認:\n"
    "  - 結果を目視で確認すること\n"
)


class GenCase(unittest.TestCase):
    """config/ と cases/ を持つ最小のプロジェクトを一時フォルダに組む。"""

    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="autotest_gen_")).resolve()
        (self.root / "config").mkdir()
        (self.root / "config" / "settings.yaml").write_text(yaml.safe_dump({
            "env": {"sandbox": True},
            "database": {"server": "unused", "database": "unused", "auth": "windows"},
            "batch": {"exe_path": "{python}", "working_dir": "."},
            "paths": {name: "sandbox/" + name for name in
                      ("input_dir", "processed_dir", "error_dir", "output_dir", "log_dir")},
        }), encoding="utf-8")
        self.md = self.root / "docs" / "cases.md"
        self.md.parent.mkdir(parents=True)

    def tearDown(self):
        shutil.rmtree(str(self.root), ignore_errors=True)

    def write_md(self, body, front=FRONT):
        self.md.write_text(front + body, encoding="utf-8")

    def gen(self, *extra):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = cli.main(["gen", "--md", str(self.md),
                             "--config", str(self.root / "config" / "settings.yaml")] + list(extra))
        return code, out.getvalue(), err.getvalue()

    def yaml_of(self, case_id, group="grp"):
        return yaml.safe_load((self.root / "cases" / group / (case_id + ".yaml")).read_text(encoding="utf-8"))

    def written(self):
        return sorted(p.relative_to(self.root).as_posix() for p in self.root.rglob("*")
                      if p.is_file() and not p.relative_to(self.root).as_posix().startswith(("config/", "docs/")))


class TestBasicGeneration(GenCase):
    def test_basic_case_maps_labels_to_yaml(self):
        self.write_md(BASIC)
        code, out, err = self.gen()
        self.assertEqual(code, 0, err)
        data = self.yaml_of("TC101")
        self.assertEqual(data["id"], "TC101")
        self.assertFalse(data["enabled"], "生成物は必ず enabled: false（未レビュー）")
        self.assertEqual(data["tags"], ["正常系"])
        self.assertEqual(data["setup"]["clean_dirs"], casegen.DEFAULT_CLEAN)
        self.assertEqual(data["execute"]["args"], ["--mode", "daily", "--date", "{date}"])
        self.assertEqual(data["assert"]["exit_code"], 0)
        self.assertEqual(data["assert"]["log"]["must_contain"], ["処理開始", "取込件数=2"])
        self.assertEqual(data["assert"]["files"][0]["exists"],
                         {"dir": "processed_dir", "pattern": "ORDER_*.csv", "count": 1})
        self.assertEqual(data["review"], ["結果を目視で確認すること"])
        self.assertIn("プロジェクト", out)

    def test_generated_case_loads_and_passes_schema(self):
        self.write_md(BASIC)
        self.assertEqual(self.gen()[0], 0)
        cases = load_cases(self.root / "cases", include_disabled=True)
        self.assertEqual([c.case_id for c in cases], ["TC101"])
        self.assertIn("grp", cases[0].tags, "出力先のフォルダ名は暗黙のタグになる")
        path = self.root / "cases" / "grp" / "TC101.yaml"
        self.assertEqual(_validate_case_schema(self.yaml_of("TC101"), path, "TC101"), [])

    def test_generated_yaml_keeps_marker_on_first_line(self):
        self.write_md(BASIC)
        self.gen()
        first = (self.root / "cases" / "grp" / "TC101.yaml").read_text(encoding="utf-8").split("\n")[0]
        self.assertTrue(first.startswith(casegen.MARKER_PREFIX))
        self.assertIn("md=docs/cases.md", first)

    def test_dry_run_writes_nothing(self):
        self.write_md(BASIC)
        code, out, _err = self.gen("--dry-run")
        self.assertEqual(code, 0)
        self.assertIn("[dry-run]", out)
        self.assertEqual(self.written(), [])

    def test_new_project_without_cases_dir_is_not_confused_with_tool_root(self):
        # cases/ がまだ無い新規プロジェクトでも、config/ の 2 階層上をルートとして扱う
        self.write_md(BASIC)
        self.assertFalse((self.root / "cases").exists())
        self.assertEqual(self.gen()[0], 0)
        self.assertTrue((self.root / "cases" / "grp" / "TC101.yaml").is_file())

    def test_kind_mapping_for_manual_and_env(self):
        self.write_md(
            "## TC201 手動：確認\n- 種別: 手動\n- 目的: 人が確認する\n- 確認: 状態を見る\n\n"
            "## TC202 環境不備：出力なし\n- 種別: 環境不備\n- 目的: 異常終了する\n"
            "- 削除: `output_dir`\n- 終了コード: 2\n")
        self.assertEqual(self.gen()[0], 0)
        manual = self.yaml_of("TC201")
        self.assertEqual(manual["mode"], "manual")
        self.assertEqual(manual["tags"], ["手動"])
        env = self.yaml_of("TC202")
        self.assertEqual(env["tags"], ["異常系", "環境不備"])
        self.assertEqual(env["setup"]["remove_dirs"], ["output_dir"])
        self.assertNotIn("output_dir", env["setup"]["clean_dirs"], "削除するフォルダはクリア対象から外す")


class TestStrictFormat(GenCase):
    """書いたものが黙って消える経路を作らない。"""

    def assert_error(self, body, *fragments):
        self.write_md(body)
        code, _out, err = self.gen()
        self.assertEqual(code, 2, "形式エラーは終了コード 2")
        for fragment in fragments:
            self.assertIn(fragment, err)
        self.assertEqual(self.written(), [], "エラーがあれば 1 ファイルも書かない")
        return err

    def test_unknown_label_is_error_with_line_number(self):
        err = self.assert_error(BASIC + "- 確認観点: typo\n", "未知の項目「確認観点」")
        self.assertRegex(err, r"L\d+ TC101")

    def test_free_text_inside_case_is_error(self):
        self.assert_error(BASIC + "これは散文です\n", "解釈できない行")

    def test_heading_inside_case_is_error(self):
        self.assert_error(BASIC + "### メモ\n", "見出しは書けません")

    def test_h2_without_space_is_error_not_silently_skipped(self):
        self.assert_error("##TC101 名前\n- 種別: 正常系\n", "空白を入れてください")

    def test_missing_required_fields(self):
        self.assert_error("## TC101 名前\n- 終了コード: 0\n", "「種別」は必須", "「目的」は必須")

    def test_case_without_any_check_is_rejected(self):
        self.assert_error("## TC101 名前\n- 種別: 正常系\n- 目的: 何か\n", "検証内容がありません")

    def test_log_words_must_be_backticked(self):
        self.assert_error(BASIC.replace("`処理開始`", "処理開始"), "`...` で囲んでください")

    def test_file_count_is_mandatory(self):
        self.assert_error(BASIC.replace(" = 1", ""), "件数は必須")

    def test_duplicate_label_is_error(self):
        self.assert_error(BASIC + "- 終了コード: 1\n", "2 回書かれています")

    def test_inline_and_children_together_is_error(self):
        self.assert_error(BASIC.replace("- 確認:\n", "- 確認: インライン\n"), "両方に書いています")

    def test_manual_case_cannot_claim_exit_code(self):
        self.assert_error("## TC101 手動\n- 種別: 手動\n- 目的: x\n- 終了コード: 0\n",
                          "手動ケースでは終了コード")

    def test_backslash_in_args_is_rejected(self):
        self.assert_error(BASIC.replace("--mode daily", r"--in C:\data"), "\\ は使えません")

    def test_unclosed_fence_is_error(self):
        self.assert_error(BASIC + "```csv input=input_dir/A.csv\nA,B\n", "閉じていません")

    def test_untyped_code_block_is_error(self):
        self.assert_error(BASIC + "```csv\nA,B\n```\n", "用途の分からないコードブロック")

    def test_yaml_block_collision_is_error_not_override(self):
        # md の「確認」と YAML ブロックの review が両方ある → どちらかが黙って消えるのを防ぐ
        self.assert_error(BASIC + "```yaml\nreview: [\"別の観点\"]\n```\n", "両方で指定されています")

    def test_yaml_block_cannot_set_generator_owned_keys(self):
        self.assert_error(BASIC + "```yaml\nenabled: true\n```\n", "YAML ブロックには書けません")

    def test_yaml_block_typo_is_caught_by_schema(self):
        self.assert_error(BASIC + "```yaml\nsetup:\n  input_files:\n    - {src: a.csv, renmae: b.csv}\n```\n",
                          "renmae")

    def test_duplicate_ids_case_insensitive(self):
        self.assert_error(BASIC + "\n" + BASIC.replace("## TC101", "## tc101"), "重複")

    def test_all_or_nothing_across_cases(self):
        good = BASIC.replace("TC101", "TC102")
        self.assert_error(good + "\n" + BASIC + "- 謎: x\n", "未知の項目")

    def test_id_too_long_for_sheet_name(self):
        self.assert_error(BASIC.replace("TC101", "T" * 32), "長すぎます")

    def test_front_matter_unknown_key(self):
        self.write_md(BASIC, front="---\n出力先: grp\n謎: 1\n---\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("フロントマターに未知の項目", err)

    def test_existing_hand_written_case_with_same_id_elsewhere(self):
        other = self.root / "cases" / "other" / "TC101.yaml"
        other.parent.mkdir(parents=True)
        other.write_text("id: TC101\nname: x\nassert: {exit_code: 0}\n", encoding="utf-8")
        self.write_md(BASIC)
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("既存のケースと重複", err)


class TestOptionsAndHelpers(GenCase):
    def test_options_override_defaults_but_not_body(self):
        self.write_md(BASIC.replace(
            "- 確認:\n", "- 回収: `output_dir/*.csv` {preview: false}\n"
                        "- DB: `T_ORDER` `SELECT ID FROM T_ORDER ORDER BY ID` {mask: [NAME]}\n- 確認:\n"))
        self.assertEqual(self.gen()[0], 0)
        data = self.yaml_of("TC101")
        self.assertEqual(data["collect"]["files"], [{"dir": "output_dir", "pattern": "*.csv", "preview": False}])
        self.assertEqual(data["snapshot"]["tables"][0]["mask"], ["NAME"])

        self.write_md(BASIC.replace("- 確認:\n", "- 回収: `output_dir/*.csv` {dir: error_dir}\n- 確認:\n"))
        code, _out, err = self.gen("--overwrite")
        self.assertEqual(code, 2)
        self.assertIn("重複", err)

    def test_clean_dirs_none_and_front_matter_default(self):
        self.write_md(BASIC.replace("- 引数:", "- クリア: なし\n- 引数:"))
        code, out, _err = self.gen()
        self.assertEqual(code, 0)
        self.assertNotIn("clean_dirs", self.yaml_of("TC101")["setup"])
        self.assertIn("clean_dirs にありません", out, "クリアしないフォルダでの件数判定は警告する")

    def test_front_matter_default_clean(self):
        self.write_md(BASIC, front="---\n出力先: grp\n既定クリア: [processed_dir]\n---\n")
        self.assertEqual(self.gen()[0], 0)
        self.assertEqual(self.yaml_of("TC101")["setup"]["clean_dirs"], ["processed_dir"])

    def test_spec_reference_goes_to_description(self):
        self.write_md(BASIC.replace("- 引数:", "- 仕様: 3.1 取込\n- 引数:"))
        self.gen()
        self.assertTrue(self.yaml_of("TC101")["description"].endswith("（仕様: 3.1 取込）"))

    def test_sql_items_inline_and_file(self):
        self.write_md(BASIC.replace("- 引数:", "- 準備SQL:\n  - `DELETE FROM T_X WHERE A = 'b'`\n"
                                              "  - `sql/common_seed.sql`\n- 引数:"))
        code, out, _err = self.gen()
        self.assertEqual(code, 0)
        self.assertEqual(self.yaml_of("TC101")["setup"]["sql"],
                         ["DELETE FROM T_X WHERE A = 'b'", {"file": "sql/common_seed.sql"}])
        self.assertIn("SQL を用意する: sql/common_seed.sql", out)

    def test_review_continuation_and_ordered_items(self):
        self.write_md(BASIC.replace("  - 結果を目視で確認すること\n",
                                    "  1. 一つ目の観点\n     続きの行\n  2. 二つ目の観点\n"))
        self.assertEqual(self.gen()[0], 0)
        self.assertEqual(self.yaml_of("TC101")["review"], ["一つ目の観点\n続きの行", "二つ目の観点"])

    def test_rules_and_comments_are_ignored(self):
        self.write_md(BASIC + "\n---\n<!-- メモ\n複数行 -->\n")
        self.assertEqual(self.gen()[0], 0)

    def test_commented_out_case_is_not_generated(self):
        # 導入部でも、ケースの後でも、コメントアウトしたケースの見出しを拾わない
        hidden = "<!--\n## TC999 無効化したケース\n- 種別: 正常系\n-->\n"
        self.write_md(hidden + BASIC + "\n" + hidden.replace("TC999", "TC998"))
        code, out, _err = self.gen()
        self.assertEqual(code, 0)
        self.assertNotIn("TC999", out)
        self.assertNotIn("TC998", out)

    def test_indented_code_block_under_bullet_closes(self):
        self.write_md(BASIC.replace("- 引数:", "- 入力: `input/A.csv` → `input_dir`\n- 引数:")
                      + "\n    ```csv file=input/A.csv\n    X,Y\n    1,2\n    ```\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 0, err)
        self.assertEqual((self.root / "cases" / "grp" / "TC101" / "input" / "A.csv").read_text(encoding="utf-8"),
                         "X,Y\n1,2\n", "ブロックの字下げは取り除いて書く")

    def test_dump_yaml_round_trips_tricky_strings(self):
        data = {"a": ["yes", "No", "null", "~", "20260803", "2026-08-03", "1.0", "a: b", "#x", "*.csv",
                      "'q'", '"dq"', "日本語", "「括弧」", "-mode", "true"],
                "b": "line1\nline2", "c": {"x": "y: z", "n": 0, "t": True}, "d": [], "e": {}}
        self.assertEqual(yaml.safe_load(casegen.dump_yaml(data)), data)


class TestDataBlocks(GenCase):
    def test_seed_generates_sql_with_null_quotes_and_identity(self):
        self.write_md(BASIC + "\n```csv seed=dbo.T_ORDER key=ID identity=on\nID,NAME,NOTE\n"
                              "1,O'Brien,(NULL)\n2,山田,\n```\n")
        code, out, _err = self.gen()
        self.assertEqual(code, 0)
        sql = (self.root / "cases" / "grp" / "TC101" / "seed.sql").read_text(encoding="utf-8")
        self.assertIn("DELETE FROM dbo.T_ORDER WHERE [ID] IN (N'1', N'2');", sql)
        self.assertIn("SET IDENTITY_INSERT dbo.T_ORDER ON;", sql)
        self.assertIn("(N'1', N'O''Brien', NULL)", sql, "(NULL) は NULL、' は二重化")
        self.assertIn("(N'2', N'山田', N'')", sql)
        self.assertIn("空欄があります", out, "空欄は数値列で 0 になるので警告する")
        self.assertEqual(self.yaml_of("TC101")["setup"]["sql"], [{"file": "cases/grp/TC101/seed.sql"}])

    def test_seed_requires_explicit_delete_policy(self):
        self.write_md(BASIC + "\n```csv seed=T_ORDER\nID\n1\n```\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("delete=all / delete=none", err)

    def test_expect_db_adds_snapshot_and_assert(self):
        self.write_md(BASIC + "\n```csv expect_db=T_ORDER key=ID ignore=UPDATED_AT fixture=after\n"
                              "ID,STATUS,UPDATED_AT\n1,完了,x\n```\n")
        self.assertEqual(self.gen()[0], 0)
        data = self.yaml_of("TC101")
        self.assertEqual(data["snapshot"]["tables"],
                         [{"name": "T_ORDER", "sql": "SELECT [ID], [STATUS], [UPDATED_AT] FROM T_ORDER ORDER BY [ID]"}])
        self.assertEqual(data["assert"]["db"], [{"table": "T_ORDER", "expected": "expected/TC101/db_T_ORDER.csv",
                                                 "key": ["ID"], "ignore_columns": ["UPDATED_AT"]}])
        self.assertTrue((self.root / "fixtures" / "TC101" / "after_T_ORDER.csv").is_file())

    def test_expect_db_key_must_exist_in_header(self):
        self.write_md(BASIC + "\n```csv expect_db=T_ORDER key=NOPE\nID\n1\n```\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("ヘッダにありません", err)

    def test_expect_file_needs_name_for_glob_and_is_written_utf8(self):
        self.write_md(BASIC + "\n```csv expect_file=output_dir/RESULT_*.csv\nA\n```\n")
        self.assertEqual(self.gen()[0], 2)
        self.write_md(BASIC + "\n```csv expect_file=output_dir/RESULT_*.csv name=R.csv encoding=cp932 "
                              "ignore_line=\"^#作成 日時\"\n名前\n山田\n```\n")
        self.assertEqual(self.gen()[0], 0)
        expected = self.root / "expected" / "TC101" / "R.csv"
        self.assertEqual(expected.read_bytes().decode("utf-8"), "名前\n山田\n",
                         "期待値は比較側が UTF-8 で読むので、encoding= に関係なく UTF-8 で書く")
        item = self.yaml_of("TC101")["assert"]["files"][-1]
        self.assertEqual(item["encoding"], "cp932")
        self.assertEqual(item["ignore_line_patterns"], ["^#作成 日時"])
        self.assertIn({"dir": "output_dir", "pattern": "RESULT_*.csv", "preview": True},
                      self.yaml_of("TC101")["collect"]["files"], "期待ファイルの実物は回収して証跡に残す")

    def test_input_block_encoding_newline_and_input_files(self):
        self.write_md(BASIC + "\n```csv input=input_dir/ORDER_1.csv encoding=cp932 newline=crlf corrupt=flip\n"
                              "名前\n山田\n```\n")
        self.assertEqual(self.gen()[0], 0)
        path = self.root / "cases" / "grp" / "TC101" / "input" / "ORDER_1.csv"
        self.assertEqual(path.read_bytes(), "名前\r\n山田\r\n".encode("cp932"))
        self.assertEqual(self.yaml_of("TC101")["setup"]["input_files"],
                         [{"src": "input/ORDER_1.csv", "dest_dir": "input_dir", "corrupt": "flip"}])

    def test_unencodable_character_is_error(self):
        self.write_md(BASIC + "\n```csv input=input_dir/A.csv encoding=cp932\n\U0001F600\n```\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("表せない文字", err)

    def test_file_block_roots(self):
        self.write_md(BASIC + "\n```text file=expected/a.txt\nE\n```\n```text file=fixtures/before_T.csv\nF\n```\n"
                              "```text file=config/app.yml\nM\n```\n")
        self.assertEqual(self.gen()[0], 0)
        for rel in ("expected/TC101/a.txt", "fixtures/TC101/before_T.csv", "cases/grp/TC101/config/app.yml"):
            self.assertTrue((self.root / rel).is_file(), rel)

    def test_unsafe_paths_are_rejected(self):
        self.write_md(BASIC + "\n```text file=../escape.txt\nX\n```\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("相対パス", err)

    def test_missing_input_file_becomes_todo(self):
        self.write_md(BASIC.replace("- 引数:", "- 入力: `input/BIG.csv` → `input_dir`\n- 引数:"))
        code, out, _err = self.gen()
        self.assertEqual(code, 0)
        self.assertIn("投入ファイルを置く: cases/grp/TC101/input/BIG.csv", out)
        self.assertTrue((self.root / "cases" / "grp" / "TC101" / "input").is_dir(), "置き場所は先に作っておく")


class TestOwnership(GenCase):
    """再生成で手作業を巻き戻さない。"""

    DATA = BASIC + "\n```csv expect_db=T_ORDER key=ID\nID,STATUS\n1,完了\n```\n"

    def yaml_path(self):
        return self.root / "cases" / "grp" / "TC101.yaml"

    def test_rerun_is_unchanged_and_md_change_needs_overwrite(self):
        self.write_md(self.DATA)
        self.assertEqual(self.gen()[0], 0)
        self.assertIn("変更なし", self.gen()[1])
        self.write_md(self.DATA.replace("目視で確認", "よく確認"))
        code, out, _err = self.gen()
        self.assertEqual(code, 1, "md が反映されていないのに成功を返さない")
        self.assertIn("差分あり", out)
        self.assertIn("[未反映]", out)
        self.assertIn("目視で確認", self.yaml_path().read_text(encoding="utf-8"), "--overwrite なしでは書き換えない")
        self.assertIn("更新", self.gen("--overwrite")[1])
        self.assertIn("よく確認", self.yaml_path().read_text(encoding="utf-8"))

    def test_hand_edited_yaml_is_never_overwritten(self):
        self.write_md(self.DATA)
        self.gen()
        text = self.yaml_path().read_text(encoding="utf-8").replace("count: 1", "count: 2")
        self.yaml_path().write_text(text, encoding="utf-8")
        self.write_md(self.DATA.replace("目視で確認", "よく確認"))
        code, out, _err = self.gen("--overwrite")
        self.assertEqual(code, 1, "手編集で md を反映できなかったのに成功を返さない")
        self.assertIn("手で編集", out)
        self.assertNotIn("--case TC101", out, "反映できなかったケースを validate / run の手順に並べない")
        self.assertIn("count: 2", self.yaml_path().read_text(encoding="utf-8"))

    def test_hand_edited_data_file_is_never_overwritten(self):
        self.write_md(self.DATA)
        self.gen()
        expected = self.root / "expected" / "TC101" / "db_T_ORDER.csv"
        expected.write_text("ID,STATUS\n1,エラー\n", encoding="utf-8")
        self.write_md(self.DATA.replace("目視で確認", "よく確認"))
        code, out, _err = self.gen("--overwrite")
        self.assertEqual(code, 1)
        self.assertIn("expected/TC101/db_T_ORDER.csv", out)
        self.assertIn("エラー", expected.read_text(encoding="utf-8"))

    def test_enabling_is_not_an_edit_but_overwrite_resets_it(self):
        self.write_md(self.DATA)
        self.gen()
        self.yaml_path().write_text(self.yaml_path().read_text(encoding="utf-8").replace(
            "enabled: false", "enabled: true"), encoding="utf-8")
        self.assertIn("変更なし", self.gen("--overwrite")[1], "レビュー済みの印を付けただけなら触らない")
        self.write_md(self.DATA.replace("目視で確認", "よく確認"))
        out = self.gen("--overwrite")[1]
        self.assertIn("enabled: false に戻しました", out)
        self.assertFalse(self.yaml_of("TC101")["enabled"], "内容が変わったら再レビューが必要")

    def test_crlf_conversion_is_not_treated_as_edit(self):
        # Windows の git（autocrlf）で改行だけ変わったファイルを手作業と誤認しない
        self.write_md(self.DATA)
        self.gen()
        for path in (self.yaml_path(), self.root / "expected" / "TC101" / "db_T_ORDER.csv"):
            path.write_bytes(path.read_bytes().replace(b"\n", b"\r\n"))
        self.assertIn("変更なし", self.gen()[1])

    def test_stale_generated_file_is_removed_on_overwrite(self):
        self.write_md(self.DATA)
        self.gen()
        stale = self.root / "expected" / "TC101" / "db_T_ORDER.csv"
        self.assertTrue(stale.is_file())
        self.write_md(BASIC)
        self.assertIn("更新", self.gen("--overwrite")[1])
        self.assertFalse(stale.exists())
        self.assertNotIn("db", self.yaml_of("TC101").get("assert", {}))

    def test_case_removed_from_md_is_reported_not_deleted(self):
        self.write_md(self.DATA + "\n" + BASIC.replace("TC101", "TC102"))
        self.gen()
        self.write_md(BASIC.replace("TC101", "TC102"))
        out = self.gen()[1]
        self.assertIn("md に無くなった生成済みケース", out)
        self.assertTrue(self.yaml_path().is_file())

    def test_md_path_with_different_case_is_same_owner(self):
        # Windows / macOS は大文字小文字を区別しない。綴りの違いで「別の md」扱いにしない
        self.write_md(self.DATA)
        self.gen()
        upper = self.root / "DOCS" / "CASES.md"
        if not upper.exists():
            self.skipTest("大文字小文字を区別するファイルシステム")
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = cli.main(["gen", "--md", str(upper), "--config", str(self.root / "config" / "settings.yaml")])
        self.assertEqual(code, 0, err.getvalue())
        self.assertIn("変更なし", out.getvalue())

    def test_renamed_md_takes_over_but_two_live_mds_conflict(self):
        self.write_md(self.DATA)
        self.gen()
        renamed = self.md.with_name("renamed.md")
        self.md.rename(renamed)
        self.md = renamed
        code, out, err = self.gen("--overwrite")
        self.assertEqual(code, 0, err)
        self.assertIn("改名・移動", out)
        self.assertIn("md=docs/renamed.md", self.yaml_path().read_text(encoding="utf-8").split("\n")[0])
        # 元の名前でもう 1 つ md を作り、同じ ID を書く → 本当の衝突
        other = renamed.with_name("cases.md")
        other.write_text(FRONT + self.DATA, encoding="utf-8")
        self.md = other
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("別の md", err)

    def test_unmarked_existing_case_is_left_alone(self):
        self.yaml_path().parent.mkdir(parents=True)
        self.yaml_path().write_text("id: TC101\nname: 手書き\nassert: {exit_code: 0}\n", encoding="utf-8")
        self.write_md(BASIC)
        code, out, _err = self.gen("--overwrite")
        self.assertEqual(code, 1)
        self.assertIn("手で管理", out)
        self.assertIn("手書き", self.yaml_path().read_text(encoding="utf-8"))


class TestDataIsNotSilentlyChanged(GenCase):
    """md に書いたデータ・検証が、エラーも出ずに変わったり消えたりしないこと（レビュー指摘の回帰）。"""

    def test_tabs_inside_blocks_are_kept(self):
        """TSV のタブが空白に化けると、投入データが黙って別物になる。"""
        self.write_md(BASIC + "\n```text input=input_dir/DATA.tsv\nA\tB\tC\n1\t2\t3\n```\n"
                              "\n- 備考: 字下げしたブロック\n\n  ```text input=input_dir/IND.txt\n  \tx\n    y\n  ```\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 0, err)
        base = self.root / "cases" / "grp" / "TC101" / "input"
        self.assertEqual((base / "DATA.tsv").read_bytes(), b"A\tB\tC\n1\t2\t3\n")
        self.assertEqual((base / "IND.txt").read_bytes(), b"\tx\n  y\n", "字下げ分だけ除き、タブは残す")

    def test_data_blocks_above_the_first_case_are_an_error(self):
        """導入部のブロックはどのケースにも入らない。黙って捨てると初期データも期待値も消える。"""
        for block in ("```csv seed=T_COMMON key=ID\nID\n1\n```", "```csv expect_db=T key=ID\nID\n1\n```",
                      "```yaml\nsetup:\n  db_lock: x\n```"):
            self.write_md(BASIC, front=FRONT + block + "\n\n")
            code, _out, err = self.gen()
            self.assertEqual(code, 2, block)
            self.assertIn("最初のケース", err)
            self.assertEqual(self.written(), [], "何も書かないこと")

    def test_plain_example_blocks_in_the_intro_are_allowed(self):
        self.write_md(BASIC, front=FRONT + "```sql\nSELECT * FROM T_ORDER;\n```\n\n")
        self.assertEqual(self.gen()[0], 0)

    def test_text_after_a_comment_on_the_same_line_is_an_error(self):
        """「<!-- 下書き --> - ログ必須: …」は行ごと読み飛ばされ、断言が消えていた。"""
        for line in ("<!-- 下書き --> - ログ禁止: `[ERROR]`\n", "<!--\n下書き\n--> - ログ禁止: `[ERROR]`\n"):
            self.write_md(BASIC + line)
            code, _out, err = self.gen()
            self.assertEqual(code, 2, line)
            self.assertIn("読み込まれません", err)

    def test_whole_line_comments_still_work(self):
        self.write_md(BASIC + "<!-- - ログ禁止: `[ERROR]` -->\n")
        self.assertEqual(self.gen()[0], 0)
        self.assertNotIn("log_must_not", self.yaml_path_text())

    def yaml_path_text(self):
        return (self.root / "cases" / "grp" / "TC101.yaml").read_text(encoding="utf-8")

    def test_space_after_comma_in_csv_is_an_error(self):
        """「1, 山田」は ' 山田' が投入される。LLM がよく書く形なので止める。"""
        for csv_text in ("ID, NAME\n1,山田", "ID,NAME\n1, 山田", "ID,NAME\n 1,山田", "ID ,NAME\n1,山田"):
            self.write_md(BASIC + "\n```csv seed=T key=ID\n%s\n```\n" % csv_text)
            code, _out, err = self.gen()
            self.assertEqual(code, 2, csv_text)
            self.assertIn("空白", err)

    def test_duplicate_keys_in_yaml_are_an_error_with_md_lines(self):
        """レビュー指摘: assert の下に log を 2 回書くと、1 回目（must_contain）が黙って消えていた。"""
        body = BASIC + "\n```yaml\nassert:\n  log:\n    must_contain: [処理開始]\n  log:\n    must_not_contain: [x]\n```\n"
        self.write_md(body)
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        lines = (FRONT + body).split("\n")
        first = lines.index("  log:") + 1
        second = lines.index("  log:", first) + 1
        self.assertIn("'log' が 2 回書かれています（L%d と L%d）" % (first, second), err)
        self.assertEqual(self.written(), [])

    def test_duplicate_keys_in_front_matter_and_options_are_an_error(self):
        self.write_md(BASIC, front="---\n出力先: grp\n出力先: other\n---\n# タイトル\n\n")
        code, _out, err = self.gen()
        self.assertEqual(code, 2)
        self.assertIn("L3", err)
        self.write_md(BASIC + "- 回収: `output_dir/*.csv` {preview: false, preview: true}\n")
        self.assertEqual(self.gen()[0], 2)

    def test_quoted_spaces_are_data(self):
        self.write_md(BASIC + '\n```csv seed=T key=ID\nID,NAME\n1," 山田 "\n```\n')
        code, _out, err = self.gen()
        self.assertEqual(code, 0, err)
        sql = (self.root / "cases" / "grp" / "TC101" / "seed.sql").read_text(encoding="utf-8")
        self.assertIn("(N'1', N' 山田 ')", sql)


class TestWritesAreAllOrNothing(GenCase):
    """書き込みの途中で失敗しても、新旧が混ざった状態を残さないこと。"""

    DATA = BASIC + ("\n```csv seed=T key=ID\nID,NAME\n1,a\n```\n"
                    "\n```csv expect_db=T key=ID\nID,NAME\n1,b\n```\n")

    def snapshot(self):
        return {rel: (self.root / rel).read_bytes() for rel in self.written()}

    def regenerate_with_failure(self, target_name, method):
        self.write_md(self.DATA)
        self.assertEqual(self.gen()[0], 0)
        before = self.snapshot()
        self.write_md(self.DATA.replace("1,a", "1,c").replace("1,b", "1,d"))
        if method == "write":
            original = Path.write_bytes

            def write_bytes(path, data):
                if path.name.startswith(target_name):
                    raise PermissionError(13, "Permission denied")
                return original(path, data)
            patch = unittest.mock.patch.object(Path, "write_bytes", write_bytes)
        else:
            original_replace = casegen.os.replace

            def replace(src, dst):
                if Path(str(src)).name == target_name:   # Excel が開いている既存ファイルの退避に失敗
                    raise PermissionError(13, "Permission denied")
                return original_replace(src, dst)
            patch = unittest.mock.patch.object(casegen.os, "replace", replace)
        with patch:
            code, _out, err = self.gen("--overwrite")
        self.assertEqual(code, 2)
        self.assertIn("書き込みエラー", err)
        self.assertIn("書く前の状態に戻しました", err)
        self.assertEqual(self.snapshot(), before, "1 ファイルでも新しい内容が残っている")
        # 記録したハッシュとデータが合ったままなので、次の gen は普通に更新できる（手編集扱いで固まらない）
        code, out, err = self.gen("--overwrite")
        self.assertEqual(code, 0, err)
        self.assertIn("更新", out)

    def test_failure_while_staging(self):
        self.regenerate_with_failure("db_T.csv", "write")

    def test_failure_while_replacing_a_locked_file(self):
        self.regenerate_with_failure("db_T.csv", "replace")

    def test_no_temporary_files_are_left(self):
        self.write_md(self.DATA)
        self.gen()
        self.assertEqual([r for r in self.written() if "autotest-gen" in r], [])


class TestEnabledSpellings(GenCase):
    def test_capitalised_true_is_reported_when_reset(self):
        """enabled: True / yes で有効にしたケースを --overwrite で戻したら、必ず知らせること。"""
        for spelling in ("True", "yes", '"true"'):
            self.write_md(BASIC)
            self.gen()
            path = self.root / "cases" / "grp" / "TC101.yaml"
            path.write_text(path.read_text(encoding="utf-8").replace("enabled: false", "enabled: %s" % spelling),
                            encoding="utf-8")
            self.write_md(BASIC.replace("目視で確認", "よく確認"))
            code, out, _err = self.gen("--overwrite")
            self.assertEqual(code, 0)
            self.assertIn("enabled: false に戻しました", out, spelling)
            shutil.rmtree(str(self.root / "cases"))


class TestUnreflectedCasesAreNotSuccess(GenCase):
    """md の変更を反映できなかったケースがあるのに、gen が成功を返して実行手順に並べないこと。"""

    def test_hand_edited_and_enabled_case_is_reported_and_excluded(self):
        self.write_md(BASIC + "\n" + BASIC.replace("TC101", "TC102"))
        self.gen()
        path = self.root / "cases" / "grp" / "TC101.yaml"
        path.write_text(path.read_text(encoding="utf-8").replace("exit_code: 0", "exit_code: 9")
                        .replace("enabled: false", "enabled: true"), encoding="utf-8")
        self.write_md((BASIC + "\n" + BASIC.replace("TC101", "TC102")).replace("- 終了コード: 0", "- 終了コード: 2"))
        code, out, _err = self.gen("--overwrite")
        self.assertEqual(code, 1)
        self.assertIn("[未反映]", out)
        self.assertIn("TC101", out.split("[未反映]")[1].splitlines()[0])
        next_steps = out.split("次の手順:")[1]
        self.assertNotIn("TC101", next_steps)
        self.assertIn("--case TC102", next_steps, "反映できたケースは手順に残す")
        code, out, _err = self.gen("--json")
        data = json.loads(out)
        self.assertEqual(code, 1)
        self.assertFalse(data["ok"])
        self.assertEqual(data["blocked"], ["TC101"])
        self.assertTrue(all("TC101" not in step for step in data["next"]))


class TestCheckAndJson(GenCase):
    def run_json(self, *extra):
        code, out, err = self.gen("--json", *extra)
        self.assertEqual(err, "")
        return code, json.loads(out)

    def test_check_reports_sync_state_without_writing(self):
        self.write_md(BASIC)
        code, out, _err = self.gen("--check")
        self.assertEqual(code, 1, "未生成は不一致")
        self.assertEqual(self.written(), [])
        self.gen()
        self.assertEqual(self.gen("--check")[0], 0)
        self.write_md(BASIC.replace("目視で確認", "よく確認"))
        code, data = self.run_json("--check")
        self.assertEqual(code, 1)
        self.assertFalse(data["in_sync"])
        self.assertEqual([(c["id"], c["action"]) for c in data["cases"]], [("TC101", "differs")])
        self.assertIn("目視で確認", (self.root / "cases" / "grp" / "TC101.yaml").read_text(encoding="utf-8"))

    def test_check_rejects_overwrite(self):
        self.write_md(BASIC)
        self.assertEqual(self.gen("--check", "--overwrite")[0], 2)

    def test_json_errors_carry_line_and_case(self):
        self.write_md(BASIC + "- 未知の項目: x\n")
        code, data = self.run_json()
        self.assertEqual(code, 2)
        self.assertFalse(data["ok"])
        self.assertEqual(data["errors"][0]["case"], "TC101")
        self.assertGreater(data["errors"][0]["line"], 0)
        self.assertIn("未知の項目", data["errors"][0]["message"])

    def test_json_success_lists_cases_and_next_steps(self):
        self.write_md(BASIC)
        code, data = self.run_json()
        self.assertEqual(code, 0)
        self.assertTrue(data["ok"])
        self.assertEqual(data["cases"][0]["action"], "new")
        self.assertIn("validate --include-disabled", data["next"][0])

    def test_check_and_json_are_gen_only(self):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = cli.main(["validate", "--json", "--config", str(self.root / "config" / "settings.yaml")])
        self.assertEqual(code, 2)


class TestDocsStayInSync(unittest.TestCase):
    """仕様書（docs）と見本（demo）が実装とずれていないこと。"""

    def test_format_doc_example_generates_without_errors(self):
        doc = (REPO_ROOT / "docs" / "ai" / "testcase-format.md").read_text(encoding="utf-8")
        example = re.search(r"````markdown\n(.*?)\n````", doc, re.S)
        self.assertIsNotNone(example, "testcase-format.md の 7 節に ````markdown の完全な例が必要")
        tmp = Path(tempfile.mkdtemp(prefix="autotest_gen_doc_")).resolve()
        try:
            md = tmp / "docs" / "example.md"
            md.parent.mkdir(parents=True)
            md.write_text(example.group(1) + "\n", encoding="utf-8")
            ctx = casegen.GenContext(tmp, tmp / "cases", md, casegen.SqlServerDialect)
            plans, _notices = casegen.plan_generation(md, ctx, overwrite=False)
            self.assertGreaterEqual(len(plans), 3)
            labels_used = set()
            for plan in plans:
                labels_used.update(plan.md.fields)
            self.assertTrue({"DB", "ファイル", "ジョブ状態", "バッチ"} <= labels_used)
        finally:
            shutil.rmtree(str(tmp), ignore_errors=True)

    def test_format_doc_lists_every_label(self):
        doc = (REPO_ROOT / "docs" / "ai" / "testcase-format.md").read_text(encoding="utf-8")
        for label in casegen.LABELS:
            self.assertIn(label, doc, "ラベル %s が testcase-format.md に書かれていない" % label)
        for attr in sorted(set().union(*casegen.BLOCK_ATTRS.values())):
            self.assertIn(attr + "=", doc, "ブロック属性 %s= が testcase-format.md に書かれていない" % attr)

    def test_md_workflow_sample_matches_generator(self):
        # 見本の生成物はコミット済み。gen の実装を変えたら見本も作り直すこと:
        #   cd demo/md_workflow && python -m autotest gen --md docs/testcases/order_import.md
        #      --config config/settings.yaml --overwrite
        root = REPO_ROOT / "demo" / "md_workflow"
        md = root / "docs" / "testcases" / "order_import.md"
        ctx = casegen.GenContext(root, root / "cases", md, casegen.SqlServerDialect)
        plans, notices = casegen.plan_generation(md, ctx, overwrite=False)
        self.assertEqual({p.action for p in plans}, {"unchanged"},
                         [(p.case_id, p.action, p.note) for p in plans if p.action != "unchanged"])
        self.assertEqual(notices, [])


if __name__ == "__main__":
    unittest.main()
