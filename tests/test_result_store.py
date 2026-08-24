# -*- coding: utf-8 -*-
"""判定結果の保存と復元（autotest report の土台）の回帰テスト。

証跡を 1 冊にまとめるとき、既存の Excel へ追記するのではなく
「保存した結果からブックを組み直す」方式を採っている。組み直しで
情報が落ちると証跡が痩せるので、往復で失われないことを固定する。

  python -m unittest discover -s tests -v
"""

import shutil
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import result_store  # noqa: E402
from autotest.config import ConfigError  # noqa: E402
from autotest.models import (NG, OK, REVIEW, CaseResult, CheckResult,  # noqa: E402
                             ImageEvidence, RunResult, Table)


def make_case_result(case_id="TC001", verdict=OK):
    result = CaseResult(case_id=case_id, name="ケース名", description="説明", tags=["正常系"])
    result.execution.command = "batch.exe --mode daily"
    result.execution.started_at = datetime(2026, 8, 7, 10, 0, 0)
    result.execution.finished_at = datetime(2026, 8, 7, 10, 0, 5, 500000)
    result.execution.exit_code = 0
    result.execution.stdout = "取込件数=2"
    result.checks.append(CheckResult("終了コード確認", "exit_code", verdict, "終了コード = 0"))
    return result


class TmpRunDir(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="autotest_store_"))
        self.run_dir = self.tmp / "20260807_100000_000"
        self.run_dir.mkdir(parents=True)

    def tearDown(self):
        shutil.rmtree(str(self.tmp), ignore_errors=True)


class TestCaseResultRoundTrip(TmpRunDir):
    def test_basic_fields_survive(self):
        original = make_case_result()
        result_store.save_case_result(original, self.run_dir)
        run = self._reload()
        got = run.cases[0]

        self.assertEqual(got.case_id, original.case_id)
        self.assertEqual(got.tags, ["正常系"])
        self.assertEqual(got.execution.exit_code, 0)
        self.assertEqual(got.execution.stdout, "取込件数=2")
        self.assertEqual(got.execution.started_at, original.execution.started_at)
        self.assertEqual(got.verdict, OK)

    def test_verdict_is_preserved_not_recomputed(self):
        """判定は採取時に確定させる。復元時に作り直さないこと。"""
        original = make_case_result(verdict=NG)
        result_store.save_case_result(original, self.run_dir)
        self.assertEqual(self._reload().cases[0].verdict, NG)

    def test_table_cell_marks_survive_tuple_key_encoding(self):
        """セルの色分け情報（タプルキー）が JSON 往復で消えないこと。"""
        original = make_case_result()
        original.db_after["T_ORDER"] = Table(
            title="T_ORDER",
            columns=["ID", "VAL"],
            rows=[["1", "a"], ["2", "b"]],
            cell_marks={(0, 1): "diff", (1, 0): "extra"},
            mono_columns=[1],
            mask_columns=["VAL"],
            note="備考",
            truncated_from=99,
        )
        result_store.save_case_result(original, self.run_dir)
        table = self._reload().cases[0].db_after["T_ORDER"]

        self.assertEqual(table.cell_marks, {(0, 1): "diff", (1, 0): "extra"})
        self.assertEqual(table.rows, [["1", "a"], ["2", "b"]])
        self.assertEqual(table.mono_columns, [1])
        self.assertEqual(table.mask_columns, ["VAL"])
        self.assertEqual(table.note, "備考")
        self.assertEqual(table.truncated_from, 99)

    def test_diff_table_inside_check_survives(self):
        original = make_case_result(verdict=NG)
        original.checks[0].diff_table = Table(
            title="差分", columns=["区分", "期待値", "実績値"],
            rows=[["相違", "100", "200"]])
        result_store.save_case_result(original, self.run_dir)
        check = self._reload().cases[0].checks[0]
        self.assertIsNotNone(check.diff_table)
        self.assertEqual(check.diff_table.rows, [["相違", "100", "200"]])

    def test_human_confirmation_survives(self):
        original = make_case_result(verdict=REVIEW)
        original.checks[0].confirmation_result = OK
        original.checks[0].confirmation_by = "山田"
        original.checks[0].confirmation_at = "2026-08-15"
        result_store.save_case_result(original, self.run_dir)
        got = self._reload().cases[0].checks[0]
        self.assertEqual(got.verdict, REVIEW, "自動判定は上書きしないこと")
        self.assertEqual(got.confirmation_result, OK)
        self.assertEqual(got.confirmation_by, "山田")
        self.assertEqual(got.confirmation_at, "2026-08-15")
        self.assertEqual(self._reload().verdict, OK)

    def test_image_paths_are_relative_so_the_folder_can_move(self):
        """証跡フォルダごと別の場所へ移しても画像を見失わないこと。"""
        evidence = self.run_dir / "evidence" / "TC001"
        evidence.mkdir(parents=True)
        image = evidence / "01_folder.png"
        image.write_bytes(b"PNG")

        original = make_case_result()
        original.images.append(ImageEvidence("フォルダ確認", image, "説明"))
        original.saved_artifacts.append(self.run_dir / "artifacts" / "TC001" / "RESULT.csv")
        result_store.save_case_result(original, self.run_dir)
        result_store.save_run_meta(self._run_meta(), self.run_dir)

        moved = self.tmp / "移動先"
        shutil.move(str(self.run_dir), str(moved))

        got = result_store.load_run_dir(moved).cases[0]
        self.assertEqual(got.images[0].path, moved / "evidence" / "TC001" / "01_folder.png")
        self.assertTrue(got.images[0].path.exists(), "移動後も画像を指していること")
        self.assertEqual(got.saved_artifacts[0], moved / "artifacts" / "TC001" / "RESULT.csv")

    def _run_meta(self, cases=None):
        run = RunResult(run_id=self.run_dir.name, started_at=datetime(2026, 8, 7, 10, 0, 0),
                        finished_at=datetime(2026, 8, 7, 10, 5, 0), env_name="結合テスト環境",
                        tester="tester")
        run.cases = cases if cases is not None else [make_case_result()]
        return run

    def _reload(self):
        result_store.save_run_meta(self._run_meta(), self.run_dir)
        return result_store.load_run_dir(self.run_dir)


class TestRunDirLoading(TmpRunDir):
    def test_missing_meta_is_reported_clearly(self):
        with self.assertRaises(ConfigError) as ctx:
            result_store.load_run_dir(self.run_dir)
        self.assertIn("結果情報がありません", str(ctx.exception))

    def test_format_version_mismatch_is_rejected(self):
        """書式が変わった結果を黙って読むと、意味の違う値で証跡を作ってしまう。"""
        results = self.run_dir / result_store.RESULTS_DIR_NAME
        results.mkdir(parents=True)
        (results / result_store.RUN_META_NAME).write_text(
            '{"format_version": 999, "run_id": "x"}', encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            result_store.load_run_dir(self.run_dir)
        self.assertIn("バージョン", str(ctx.exception))

    def test_case_order_is_preserved(self):
        run = RunResult(run_id="r", started_at=datetime(2026, 8, 7, 10, 0, 0))
        for case_id in ("TC003", "TC001", "TC002"):
            result = make_case_result(case_id)
            run.cases.append(result)
            result_store.save_case_result(result, self.run_dir)
        result_store.save_run_meta(run, self.run_dir)

        got = [c.case_id for c in result_store.load_run_dir(self.run_dir).cases]
        self.assertEqual(got, ["TC003", "TC001", "TC002"], "実行順が保たれること")

    def test_manual_pending_survives(self):
        run = RunResult(run_id="r", started_at=datetime(2026, 8, 7, 10, 0, 0))
        run.manual_pending = ["TC008_manual"]
        run.cases.append(make_case_result())
        result_store.save_case_result(run.cases[0], self.run_dir)
        result_store.save_run_meta(run, self.run_dir)
        self.assertEqual(result_store.load_run_dir(self.run_dir).manual_pending, ["TC008_manual"])


class TestMergeRuns(unittest.TestCase):
    def _run(self, run_id, case_ids, manual_pending=None, env="環境A"):
        run = RunResult(run_id=run_id, started_at=datetime(2026, 8, 7, 10, 0, 0),
                        finished_at=datetime(2026, 8, 7, 10, 5, 0), env_name=env)
        run.cases = [make_case_result(cid) for cid in case_ids]
        run.manual_pending = list(manual_pending or [])
        return run

    def test_cases_from_all_sources_are_included(self):
        merged = result_store.merge_runs(
            [self._run("r1", ["TC001", "TC002"]), self._run("r2", ["TC008"])], ["r1", "r2"])
        self.assertEqual([c.case_id for c in merged.cases], ["TC001", "TC002", "TC008"])
        self.assertIn("統合レポート", merged.filter_description)

    def test_duplicate_case_across_sources_is_rejected(self):
        """どちらの結果を採用したか分からない証跡は成立しない。"""
        with self.assertRaises(ConfigError) as ctx:
            result_store.merge_runs(
                [self._run("r1", ["TC001"]), self._run("r2", ["TC001"])], ["r1", "r2"])
        message = str(ctx.exception)
        self.assertIn("TC001", message)
        self.assertIn("r1", message)
        self.assertIn("r2", message)

    def test_collected_manual_case_leaves_the_pending_list(self):
        merged = result_store.merge_runs(
            [self._run("r1", ["TC001"], manual_pending=["TC008"]),
             self._run("manual_TC008", ["TC008"])],
            ["r1", "manual_TC008"])
        self.assertEqual(merged.manual_pending, [], "採取済みの手動ケースは未実施から外れる")

    def test_uncollected_manual_case_stays_pending(self):
        merged = result_store.merge_runs(
            [self._run("r1", ["TC001"], manual_pending=["TC008", "TC009"]),
             self._run("manual_TC008", ["TC008"])],
            ["r1", "manual_TC008"])
        self.assertEqual(merged.manual_pending, ["TC009"])
        self.assertEqual(merged.total_available, 3, "未採取分も全体件数に数える")
        self.assertEqual(merged.verdict, REVIEW, "未採取が残るのに合格を返さないこと")
        self.assertEqual(merged.review_count, 1)

    def test_mixed_environments_are_flagged(self):
        """別環境の結果を 1 冊にすると、どの環境の結果か分からなくなる。"""
        merged = result_store.merge_runs(
            [self._run("r1", ["TC001"], env="結合テスト環境"),
             self._run("r2", ["TC002"], env="本番相当環境")], ["r1", "r2"])
        self.assertIn("混在", merged.env_name)

    def test_verdict_reflects_worst_case(self):
        ng = self._run("r2", [])
        ng.cases = [make_case_result("TC900", verdict=NG)]
        merged = result_store.merge_runs([self._run("r1", ["TC001"]), ng], ["r1", "r2"])
        self.assertEqual(merged.verdict, NG)

    def test_review_case_keeps_the_report_from_passing(self):
        review = self._run("manual_TC008", [])
        review.cases = [make_case_result("TC008", verdict=REVIEW)]
        merged = result_store.merge_runs([self._run("r1", ["TC001"]), review], ["r1", "m"])
        self.assertEqual(merged.verdict, REVIEW, "確認待ちが残る間は合格にしない")

    def test_empty_input_is_rejected(self):
        with self.assertRaises(ConfigError):
            result_store.merge_runs([], [])


if __name__ == "__main__":
    unittest.main()
