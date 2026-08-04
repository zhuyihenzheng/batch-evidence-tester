# -*- coding: utf-8 -*-
"""レビュー指摘の回帰テスト。

このツールは「他のプログラムが正しいか」を判定する立場なので、
偽 OK（本当は NG なのに OK と出る）は最も重い欠陥として扱う。
以下はすべて偽 OK か破壊的操作に関する再現テスト。

標準ライブラリの unittest だけを使う（Anaconda 5.2 に追加インストール不要）。

  python -m unittest discover -s tests -v
"""

import os
import shutil
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import compare, fsops, logs  # noqa: E402
from autotest.config import ConfigError, Settings, load_cases  # noqa: E402
from autotest.models import NG, OK, SKIP, CaseResult, RunResult, Table  # noqa: E402


def make_table(title, columns, rows, **kw):
    return Table(title=title, columns=columns, rows=rows, **kw)


class TmpDirCase(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="autotest_ut_"))

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write_settings(self, paths, extra=None):
        """最小構成の Settings を組み立てる。"""
        raw = {
            "batch": {"exe_path": "dummy.exe"},
            "database": {"server": "s", "database": "d"},
            "paths": paths,
            "log": {"patterns": ["*.log"], "encoding": "utf-8"},
        }
        if extra:
            raw.update(extra)
        return Settings(raw=raw, source=self.tmp / "settings.yaml", project_root=self.tmp)


# =============================================================================
# 1. DB スナップショットの打ち切りが判定に影響してはならない
# =============================================================================
class TestDbTruncationNotUsedForJudgment(unittest.TestCase):
    def test_row_beyond_display_limit_is_still_compared(self):
        """表示上限を超えた行の相違も NG として検出されること。

        表示用の打ち切りと判定用データが同一だと、上限より後ろの異常が
        永久に検査されない（偽 OK）。
        """
        columns = ["ID", "VAL"]
        expected_rows = [[str(i), "ok"] for i in range(1, 502)]
        actual_rows = [[str(i), "ok"] for i in range(1, 501)] + [["501", "BROKEN"]]

        result = compare.compare_db_table(
            "T", make_table("T", columns, actual_rows),
            make_table("T", columns, expected_rows),
            keys=["ID"],
        )
        self.assertEqual(result.verdict, NG, "501 行目の相違が検出されていない（偽 OK）")

    def test_truncated_table_is_rejected_for_comparison(self):
        """打ち切り済みの Table を判定に使ったら明示的に NG にすること。"""
        columns = ["ID", "VAL"]
        actual = make_table("T", columns, [["1", "ok"]], truncated_from=999)
        expected = make_table("T", columns, [["1", "ok"]])

        result = compare.compare_db_table("T", actual, expected, keys=["ID"])
        self.assertEqual(result.verdict, NG, "打ち切られたデータで OK 判定してはいけない")
        self.assertIn("打ち切", result.detail)


# =============================================================================
# 6. DB の主キー重複を黙って握りつぶしてはならない
# =============================================================================
class TestDuplicateKeyDetection(unittest.TestCase):
    def test_duplicate_key_in_actual_is_ng(self):
        """実績側に同一キーが 2 件あり片方が異常なら NG。

        dict 索引で後勝ちにすると、正しい方だけが残って OK になる（偽 OK）。
        """
        columns = ["ID", "VAL"]
        actual = make_table("T", columns, [["1", "BROKEN"], ["1", "ok"]])
        expected = make_table("T", columns, [["1", "ok"]])

        result = compare.compare_db_table("T", actual, expected, keys=["ID"])
        self.assertEqual(result.verdict, NG, "キー重複が見逃されている（偽 OK）")
        self.assertIn("重複", result.detail)

    def test_duplicate_key_in_expected_is_ng(self):
        columns = ["ID", "VAL"]
        actual = make_table("T", columns, [["1", "ok"]])
        expected = make_table("T", columns, [["1", "ok"], ["1", "ok"]])

        result = compare.compare_db_table("T", actual, expected, keys=["ID"])
        self.assertEqual(result.verdict, NG)


# =============================================================================
# 2. ログの表示打ち切りがキーワード判定に影響してはならない
# =============================================================================
class TestLogTruncationNotUsedForAssertion(TmpDirCase):
    def _slice_with(self, lines):
        return logs.LogSlice(path=self.tmp / "a.log", lines=lines, method="offset",
                             total_lines_before_trim=len(lines))

    def test_error_in_middle_of_long_log_is_detected(self):
        """中略された範囲にある [ERROR] も禁止キーワードとして検出されること。"""
        log_dir = self.tmp / "log"
        log_dir.mkdir()
        lines = ["line %d" % i for i in range(1000)]
        lines[500] = "2026-08-04 10:00:00 [ERROR] 想定外エラー"
        (log_dir / "a.log").write_text("\n".join(lines), encoding="utf-8")

        settings = self.write_settings({"log_dir": str(log_dir)},
                                       extra={"log": {"patterns": ["*.log"], "encoding": "utf-8",
                                                      "max_lines_in_excel": 10}})
        offsets = {}  # 実行前にファイルが無かった＝全文読み
        slices = logs.collect(settings, offsets, datetime.now(), datetime.now(), log_dir=log_dir)

        _missing, forbidden = logs.check_keywords(slices, [], ["[ERROR]"])
        self.assertEqual(forbidden, ["[ERROR]"], "中略部分の ERROR が見逃されている（偽 OK）")

    def test_rotated_file_is_read_from_head(self):
        """ローテートで中身が入れ替わったら、旧オフセットではなく先頭から読むこと。"""
        log_dir = self.tmp / "log"
        log_dir.mkdir()
        path = log_dir / "a.log"

        path.write_text("OLD-A\nOLD-B\nOLD-C\n", encoding="utf-8")
        settings = self.write_settings({"log_dir": str(log_dir)})
        offsets = logs.snapshot_offsets(settings, log_dir)

        # 別内容・かつ旧サイズより大きいファイルに差し替える
        path.write_text("NEW-1 [ERROR] 先頭行\nNEW-2\nNEW-3\nNEW-4\nNEW-5\n", encoding="utf-8")

        slices = logs.collect(settings, offsets, datetime.now(), datetime.now(), log_dir=log_dir)
        text = "\n".join(s.text for s in slices)
        self.assertIn("NEW-1", text, "差し替え後の先頭行が読めていない")


# =============================================================================
# 5. クリア対象の実パスに危険なものを許してはならない
# =============================================================================
class TestClearDirSafety(TmpDirCase):
    """危険パスの拒否テスト。

    ！重要！ここでは実在の危険パス（/ や ~）を clear_dir に渡してはいけない。
    ガードが未実装・退行した瞬間に本物のファイルが消える。
    そのため「判定関数 assert_safe_to_clear() だけを呼ぶ」設計にし、
    削除を伴う経路は一時ディレクトリ配下でしか実行しない。
    """

    def test_rejects_dangerous_paths(self):
        dangerous = [
            Path(os.path.abspath(os.sep)),   # ファイルシステムルート
            Path.home(),                     # ユーザーホーム
            Path.home().parent,              # /Users
            self.tmp,                        # プロジェクトルートそのもの
            self.tmp.parent,                 # プロジェクトルートの上位
        ]
        for target in dangerous:
            with self.subTest(target=str(target)):
                with self.assertRaises(ConfigError):
                    # 削除は一切行わない。判定関数のみを検証する
                    fsops.assert_safe_to_clear(target, project_root=self.tmp)

    def test_allows_normal_subdirectory(self):
        target = self.tmp / "work" / "in"
        target.mkdir(parents=True)
        fsops.assert_safe_to_clear(target, project_root=self.tmp)  # 例外が出なければ OK

    def test_clear_dir_rejects_dangerous_alias_without_deleting(self):
        """clear_dir 経由でも、危険パスなら 1 件も消さずに例外を投げること。"""
        canary = self.tmp / "canary.txt"
        canary.write_text("must survive", encoding="utf-8")

        settings = self.write_settings({"input_dir": str(self.tmp)})  # = プロジェクトルート
        with self.assertRaises(ConfigError):
            fsops.clear_dir(settings, "input_dir")
        self.assertTrue(canary.exists(), "危険判定なのにファイルが削除された")

    def test_clears_normal_subdirectory(self):
        target = self.tmp / "work" / "in"
        target.mkdir(parents=True)
        (target / "x.txt").write_text("x", encoding="utf-8")
        (target / "sub").mkdir()
        (target / "sub" / "y.txt").write_text("y", encoding="utf-8")

        settings = self.write_settings({"input_dir": str(target)})
        self.assertEqual(fsops.clear_dir(settings, "input_dir"), 2)
        self.assertEqual(list(target.iterdir()), [])


# =============================================================================
# 7. 実行ケースがゼロなら成功にしてはならない
# =============================================================================
class TestZeroCases(TmpDirCase):
    def test_all_disabled_raises(self):
        cases_dir = self.tmp / "cases"
        cases_dir.mkdir()
        (cases_dir / "a.yaml").write_text(
            "id: A\nname: A\nenabled: false\n", encoding="utf-8")
        with self.assertRaises(ConfigError):
            load_cases(cases_dir)

    def test_run_result_with_no_cases_is_not_ok(self):
        run = RunResult(run_id="x", started_at=datetime.now())
        self.assertNotEqual(run.verdict, OK, "0 ケースで OK は偽グリーン")

    def test_run_result_all_skip_is_not_ok(self):
        run = RunResult(run_id="x", started_at=datetime.now())
        case = CaseResult(case_id="A", name="A")
        run.cases.append(case)  # checks 空 = SKIP
        self.assertNotEqual(run.verdict, OK)
        self.assertEqual(case.verdict, SKIP)


# =============================================================================
# 10. ケース ID 重複は設定エラー
# =============================================================================
class TestDuplicateCaseId(TmpDirCase):
    def test_duplicate_id_rejected(self):
        cases_dir = self.tmp / "cases"
        cases_dir.mkdir()
        for name in ("a.yaml", "b.yaml"):
            (cases_dir / name).write_text("id: DUP\nname: %s\n" % name, encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            load_cases(cases_dir)
        self.assertIn("DUP", str(ctx.exception))


# =============================================================================
# 8. Excel 生成が内容で壊れたり、数式として解釈されたりしてはならない
# =============================================================================
class TestExcelContentSafety(TmpDirCase):
    def _build(self, rows):
        from autotest.excel import build_workbook

        run = RunResult(run_id="ut", started_at=datetime.now(), finished_at=datetime.now())
        case = CaseResult(case_id="TC", name="TC")
        case.db_after["T"] = make_table("T", ["ID", "VAL"], rows)
        case.checks.append(compare.CheckResult("dummy", "db", OK, "ok"))
        run.cases.append(case)
        out = self.tmp / "e.xlsx"
        build_workbook(run, out, {"max_db_rows": 500})
        return out

    def test_control_character_does_not_break_workbook(self):
        """NUL 等の制御文字が混ざっても証跡簿の生成が失敗しないこと。"""
        out = self._build([["1", "bad\x00value\x07"]])
        self.assertTrue(out.exists())

    def test_formula_like_value_is_not_written_as_formula(self):
        """`=1+1` のような値が Excel 数式として解釈されないこと。"""
        from openpyxl import load_workbook

        out = self._build([["1", "=1+1"], ["2", "+SUM(A1)"], ["3", "-2"], ["4", "@x"]])
        ws = load_workbook(out)["TC"]
        formula_cells = [c.coordinate for row in ws.iter_rows() for c in row if c.data_type == "f"]
        self.assertEqual(formula_cells, [], "数式セルとして書かれている（インジェクション）")
        texts = [c.value for row in ws.iter_rows() for c in row
                 if isinstance(c.value, str) and "1+1" in c.value]
        self.assertTrue(texts, "値そのものが失われている")


# =============================================================================
# 9. 出力ファイルの glob が複数ヒットしたら曖昧として扱う
# =============================================================================
class TestAmbiguousFileMatch(TmpDirCase):
    def test_multiple_matches_reported(self):
        d = self.tmp / "out"
        d.mkdir()
        (d / "RESULT_1.csv").write_text("a", encoding="utf-8")
        (d / "RESULT_2.csv").write_text("b", encoding="utf-8")
        self.assertEqual(len(fsops.find_files(d, "RESULT_*.csv")), 2)


# =============================================================================
# 11. compare_text_file は actual_path だけでも比較できること
# =============================================================================
class TestCompareTextFileFromPath(TmpDirCase):
    def test_identical_file_is_ok_without_actual_text(self):
        expected = self.tmp / "exp.csv"
        actual = self.tmp / "act.csv"
        expected.write_text("a,b\n1,2\n", encoding="utf-8")
        actual.write_text("a,b\n1,2\n", encoding="utf-8")

        result = compare.compare_text_file("f", actual, expected)
        self.assertEqual(result.verdict, OK, "同一内容なのに NG になっている")


if __name__ == "__main__":
    unittest.main(verbosity=2)
