# -*- coding: utf-8 -*-
"""手動実施ケース（mode: manual）の回帰テスト。

このモードは「人が手で batch を動かす」ものなので、自動判定の一致は
参考値でしかない。ここで OK を通すと「人が確認していないのに合格」が
成立してしまう —— 偽 OK の中でも最も危ないかたち。
その一線と、before / after を跨ぐ状態の引き継ぎを固定する。

  python -m unittest discover -s tests -v
"""

import shutil
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import manual as manual_mod  # noqa: E402
from autotest.config import ConfigError, Settings, TestCase, load_cases  # noqa: E402
from autotest.models import OK, REVIEW, SKIP  # noqa: E402


class TmpCase(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="autotest_manual_"))

    def tearDown(self):
        shutil.rmtree(str(self.tmp), ignore_errors=True)

    def write_settings(self, paths=None, extra=None):
        raw = {
            "batch": {"exe_path": "dummy.exe"},
            "database": {"server": "s", "database": "d"},
            "paths": paths if paths is not None else {"log_dir": str(self.tmp / "log")},
            "log": {"patterns": ["*.log"], "encoding": "utf-8"},
        }
        if extra:
            raw.update(extra)
        return Settings(raw=raw, source=self.tmp / "settings.yaml", project_root=self.tmp)


# =============================================================================
# ケース定義（mode キー）
# =============================================================================
class TestManualModeKey(TmpCase):
    def _load(self, body):
        root = self.tmp / "cases"
        root.mkdir(parents=True, exist_ok=True)
        (root / "TC_M.yaml").write_text(body, encoding="utf-8")
        return load_cases(root)[0]

    def test_mode_defaults_to_auto(self):
        case = self._load("id: TC_M\nname: M\n")
        self.assertEqual(case.mode, "auto")
        self.assertFalse(case.is_manual)

    def test_manual_mode_is_recognised(self):
        case = self._load("id: TC_M\nname: M\nmode: manual\n")
        self.assertTrue(case.is_manual)

    def test_unknown_mode_is_rejected(self):
        """綴り間違いを黙って auto 扱いにすると、手動のつもりが自動実行される。"""
        with self.assertRaises(ConfigError) as ctx:
            self._load("id: TC_M\nname: M\nmode: manaul\n")
        self.assertIn("mode", str(ctx.exception))


# =============================================================================
# preflight（手動では戻し役がいない機能を禁じる）
# =============================================================================
class TestManualPreflight(TmpCase):
    def _problems(self, setup):
        from autotest.orchestrator import preflight_case

        exe = self.tmp / "b.exe"
        exe.write_text("x", encoding="utf-8")
        settings = self.write_settings(extra={"batch": {"exe_path": str(exe)}})
        case = TestCase(case_id="TC_M", name="M", source=self.tmp / "TC_M.yaml",
                        mode="manual", setup=setup)
        return preflight_case(settings, case)

    def test_replace_files_is_refused_for_manual(self):
        """差し替えの復元は自動実行の finally 依存。2 プロセスでは担保できない。"""
        problems = self._problems({"replace_files": [{"src": "a.config", "dest_dir": "log_dir"}]})
        self.assertTrue(any("replace_files" in p for p in problems), problems)

    def test_db_lock_is_refused_for_manual(self):
        problems = self._problems({"db_lock": "BEGIN TRAN; UPDATE T SET A = A"})
        self.assertTrue(any("db_lock" in p for p in problems), problems)

    def test_auto_case_still_allows_them(self):
        from autotest.orchestrator import preflight_case

        exe = self.tmp / "b.exe"
        exe.write_text("x", encoding="utf-8")
        settings = self.write_settings(extra={"batch": {"exe_path": str(exe)}})
        case = TestCase(case_id="TC_A", name="A", source=self.tmp / "TC_A.yaml",
                        setup={"db_lock": "BEGIN TRAN"})
        self.assertEqual([p for p in preflight_case(settings, case) if "db_lock" in p], [])


# =============================================================================
# session.json（before と after を跨ぐ状態）
# =============================================================================
class TestManualSession(TmpCase):
    def _session(self, **kw):
        params = dict(
            session_id="manual_TC_M_20260807_100000_000",
            case_id="TC_M",
            base_date="20260807",
            log_dir=str(self.tmp / "log"),
            before_started_at=datetime(2026, 8, 7, 10, 0, 0),
            before_finished_at=datetime(2026, 8, 7, 10, 0, 5),
        )
        params.update(kw)
        return manual_mod.ManualSession(**params)

    def test_round_trip(self):
        session_dir = self.tmp / "session"
        original = self._session(offline=True, batch_name="invoice",
                                 snapshot_tables=["T_ORDER"])
        original.save(session_dir)
        got = manual_mod.ManualSession.load(session_dir)

        self.assertEqual(got.case_id, "TC_M")
        self.assertEqual(got.base_date, "20260807")
        self.assertTrue(got.offline)
        self.assertEqual(got.batch_name, "invoice")
        self.assertEqual(got.snapshot_tables, ["T_ORDER"])
        self.assertEqual(got.before_finished_at, datetime(2026, 8, 7, 10, 0, 5))

    def test_log_offsets_round_trip(self):
        """ログの読み取り位置が化けると、増分ではなく全文を貼ることになる。"""
        offsets = {Path("/var/log/batch.log"): (10432, 4096, "abc123")}
        session_dir = self.tmp / "session"
        self._session(log_offsets=manual_mod.ManualSession.encode_offsets(offsets)).save(session_dir)

        restored = manual_mod.ManualSession.load(session_dir).decode_offsets()
        self.assertEqual(restored, offsets)

    def test_missing_file_is_reported_with_guidance(self):
        with self.assertRaises(ConfigError) as ctx:
            manual_mod.ManualSession.load(self.tmp / "無い")
        self.assertIn("--phase before", str(ctx.exception))

    def test_format_version_mismatch_is_rejected(self):
        session_dir = self.tmp / "session"
        session_dir.mkdir()
        (session_dir / "session.json").write_text(
            '{"format_version": 999, "session_id": "x", "case_id": "y",'
            ' "base_date": "20260807", "phase": "before_done"}', encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            manual_mod.ManualSession.load(session_dir)
        self.assertIn("バージョン", str(ctx.exception))

    def test_missing_required_key_is_rejected(self):
        """欠けた項目を既定値で埋めると、当日日付で {date} を展開して静かにずれる。"""
        session_dir = self.tmp / "session"
        session_dir.mkdir()
        (session_dir / "session.json").write_text(
            '{"format_version": 1, "session_id": "x", "case_id": "y", "phase": "before_done"}',
            encoding="utf-8")
        with self.assertRaises(ConfigError) as ctx:
            manual_mod.ManualSession.load(session_dir)
        self.assertIn("base_date", str(ctx.exception))

    def test_find_sessions_lists_only_open_ones(self):
        out = self.tmp / "output"
        for name, phase in (("manual_TC_M_20260807_100000_000", manual_mod.PHASE_BEFORE_DONE),
                            ("manual_TC_M_20260807_110000_000", manual_mod.PHASE_FINISHED),
                            ("manual_TC_M_20260807_120000_000", manual_mod.PHASE_BEFORE_DONE)):
            self._session(session_id=name, phase=phase).save(out / name)

        names = [p.name for p in manual_mod.find_sessions(out, "TC_M")]
        self.assertEqual(names, ["manual_TC_M_20260807_120000_000",
                                 "manual_TC_M_20260807_100000_000"],
                         "採取済みは除外し、新しい順に返すこと")

    def test_before_images_survive_the_process_boundary(self):
        """実行前に撮った画像が after 側の証跡に入ること。

        before と after は別プロセスなので、画像の表題・説明はメモリでは
        渡らない。ここが抜けると証跡から「実行前」が丸ごと消え、
        手動実施モードの意味（実行前後の対比）が失われる。
        """
        from autotest.models import ImageEvidence

        session_dir = self.tmp / "session"
        evidence = session_dir / "evidence" / "TC_M"
        evidence.mkdir(parents=True)
        image = evidence / "01_folder_input_dir_実行前.png"
        image.write_bytes(b"PNG")

        images = [ImageEvidence("フォルダ確認: input_dir（実行前）", image, "3 件")]
        session = self._session(
            before_images=manual_mod.ManualSession.encode_images(images, session_dir))
        session.save(session_dir)

        restored = manual_mod.ManualSession.load(session_dir).decode_images(session_dir)
        self.assertEqual(len(restored), 1)
        self.assertEqual(restored[0].title, "フォルダ確認: input_dir（実行前）")
        self.assertEqual(restored[0].caption, "3 件")
        self.assertEqual(restored[0].path, image)
        self.assertTrue(restored[0].path.exists())

    def test_before_images_are_stored_relative_so_the_folder_can_move(self):
        from autotest.models import ImageEvidence

        session_dir = self.tmp / "session"
        image = session_dir / "evidence" / "TC_M" / "01_x.png"
        encoded = manual_mod.ManualSession.encode_images(
            [ImageEvidence("t", image, "c")], session_dir)
        self.assertEqual(encoded[0]["path"], "evidence/TC_M/01_x.png")

    def test_find_sessions_ignores_other_cases(self):
        out = self.tmp / "output"
        self._session(session_id="manual_TC_OTHER_20260807_100000_000",
                      case_id="TC_OTHER").save(out / "manual_TC_OTHER_20260807_100000_000")
        self.assertEqual(manual_mod.find_sessions(out, "TC_M"), [])


# =============================================================================
# 判定（手動ケースは自動では合格しない）
# =============================================================================
class TestManualVerdict(TmpCase):
    """run_manual_after の判定変換。ここが崩れると人の確認なしで合格が出る。"""

    def _run_after(self, assertions, log_lines="処理正常終了\n"):
        from autotest import db as dbm
        from autotest.orchestrator import CaseRunner

        log_dir = self.tmp / "log"
        log_dir.mkdir(parents=True, exist_ok=True)
        (log_dir / "batch.log").write_text(log_lines, encoding="utf-8")

        exe = self.tmp / "b.exe"
        exe.write_text("x", encoding="utf-8")
        settings = self.write_settings(
            paths={"log_dir": str(log_dir), "output_dir": str(self.tmp / "out")},
            extra={"batch": {"exe_path": str(exe)}})

        case = TestCase(case_id="TC_M", name="M", source=self.tmp / "TC_M.yaml",
                        mode="manual", assertions=assertions)
        session_dir = self.tmp / "session"
        session_dir.mkdir(parents=True, exist_ok=True)
        session = manual_mod.ManualSession(
            session_id="manual_TC_M_1", case_id="TC_M", base_date="20260807",
            log_dir=str(log_dir),
            before_started_at=datetime(2026, 8, 7, 10, 0, 0),
            before_finished_at=datetime(2026, 8, 7, 10, 0, 5))

        class FakeClient(object):
            def query(self, sql, params=None):
                return [], []

            def execute_script(self, sql):
                pass

            def close(self):
                pass

        saved = dbm.create_client
        try:
            dbm.create_client = lambda *a, **k: FakeClient()
            runner = CaseRunner(settings, session_dir)
            return runner.run_manual_after(case, session, session_dir)
        finally:
            dbm.create_client = saved

    def test_all_passing_case_is_still_only_review(self):
        """自動比較が全部一致しても OK にはしない。"""
        result = self._run_after({"log": {"must_contain": ["処理正常終了"]}})
        self.assertEqual(result.verdict, REVIEW)
        self.assertNotEqual(result.verdict, OK)
        self.assertTrue(any(c.verdict == REVIEW for c in result.checks))

    def test_banner_check_is_always_present(self):
        """判定項目が 1 つも無くても「確認待ち」で残ること（SKIP で流さない）。"""
        result = self._run_after({})
        self.assertEqual(result.checks[0].category, "manual")
        self.assertEqual(result.verdict, REVIEW)

    def test_failing_check_stays_ng(self):
        """NG は確認を待たずに問題として扱う。"""
        result = self._run_after({"log": {"must_contain": ["出るはずのない文字列"]}})
        self.assertEqual(result.verdict, "NG")

    def test_exit_code_assertion_is_skipped_not_judged(self):
        """手動起動では終了コードを取得できない。自己申告で合格を出させない。"""
        result = self._run_after({"exit_code": 0})
        exit_checks = [c for c in result.checks if c.category == "exit_code"]
        self.assertEqual(len(exit_checks), 1)
        self.assertEqual(exit_checks[0].verdict, SKIP)
        self.assertIn("取得していません", exit_checks[0].detail)

    def test_base_date_comes_from_session_not_today(self):
        """before と after が日をまたいでも {date} の展開がずれないこと。"""
        from autotest import db as dbm
        from autotest.orchestrator import CaseRunner

        log_dir = self.tmp / "log"
        log_dir.mkdir(parents=True, exist_ok=True)
        exe = self.tmp / "b.exe"
        exe.write_text("x", encoding="utf-8")
        settings = self.write_settings(paths={"log_dir": str(log_dir)},
                                       extra={"batch": {"exe_path": str(exe)}})
        case = TestCase(case_id="TC_M", name="M", source=self.tmp / "TC_M.yaml", mode="manual")
        session = manual_mod.ManualSession(
            session_id="s", case_id="TC_M", base_date="20200101", log_dir=str(log_dir),
            before_finished_at=datetime(2020, 1, 1, 10, 0, 0))

        class FakeClient(object):
            def query(self, sql, params=None):
                return [], []

            def execute_script(self, sql):
                pass

            def close(self):
                pass

        saved = dbm.create_client
        try:
            dbm.create_client = lambda *a, **k: FakeClient()
            CaseRunner(settings, self.tmp / "s").run_manual_after(case, session, self.tmp / "s")
        finally:
            dbm.create_client = saved

        self.assertEqual(settings.base_date, date(2020, 1, 1),
                         "当日ではなく session の基準日を使うこと")


if __name__ == "__main__":
    unittest.main()
