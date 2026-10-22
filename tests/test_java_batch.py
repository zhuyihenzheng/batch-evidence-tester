# -*- coding: utf-8 -*-
"""Java（Spring Boot）batch の起動と、ジョブ状態判定（assert.job_status）の回帰テスト。

Spring Boot は作りによって、ジョブが FAILED でも終了コード 0 で終わる。
終了コードだけで合否を決めると偽 OK になるため、ジョブランチャの完了行を読む
job_status を用意した。ここではその「NG に倒れるべき場面で NG になる」ことを守る。

  python -m unittest tests.test_java_batch -v
"""

import io
import os
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import cli, runner  # noqa: E402
from autotest.config import ConfigError, _validate_case_schema, load_cases, load_settings  # noqa: E402
from autotest.logs import LogSlice  # noqa: E402
from autotest.models import NG, OK, REVIEW, ExecutionInfo  # noqa: E402
from autotest.orchestrator import CaseRunner, _assert_job_status, preflight_case  # noqa: E402

REPO_ROOT = Path(__file__).resolve().parent.parent
FAKE = REPO_ROOT / "demo" / "fake_spring_batch.py"


def completed(job, status, params="{'targetDate':'{value=20260803, type=class java.lang.String, identifying=true}'}"):
    return ("2026-08-03T10:00:00.000+09:00  INFO 1 --- [main] o.s.b.c.l.s.TaskExecutorJobLauncher : "
            "Job: [SimpleJob: [name=%s]] completed with the following parameters: [{%s}] "
            "and the following status: [%s] in 35ms" % (job, params, status))


def launched(job, params):
    return ("2026-08-03T10:00:00.000+09:00  INFO 1 --- [main] o.s.b.c.l.s.TaskExecutorJobLauncher : "
            "Job: [SimpleJob: [name=%s]] launched with the following parameters: [%s]" % (job, params))


def check(spec, stdout="", logs=()):
    slices = [LogSlice(Path("x.log"), list(text.split("\n")), "offset") for text in logs]
    return _assert_job_status(spec, ExecutionInfo(stdout=stdout), slices)


class TestJobStatus(unittest.TestCase):
    def test_completed_matches(self):
        result = check("COMPLETED", stdout=completed("orderJob", "COMPLETED"))
        self.assertEqual(result.verdict, OK)
        self.assertIn("orderJob=COMPLETED", result.detail)

    def test_failed_job_is_ng_even_if_expected_completed(self):
        result = check("COMPLETED", stdout=completed("orderJob", "FAILED"))
        self.assertEqual(result.verdict, NG)
        self.assertIn("FAILED", result.detail)

    def test_no_completion_line_is_ng(self):
        # 成功の証拠が無い = NG。ログレベルを絞って完了行が出ない構成でも OK にしない
        result = check("COMPLETED", stdout="Started OrderBatchApplication in 1.2 seconds")
        self.assertEqual(result.verdict, NG)
        self.assertIn("INFO 未満", result.detail)

    def test_already_complete_instance_is_explained(self):
        result = check("COMPLETED", stdout="JobInstanceAlreadyCompleteException: A job instance already "
                                           "exists and is complete for identifying parameters={}")
        self.assertEqual(result.verdict, NG)
        self.assertIn("RunIdIncrementer", result.detail)

    def test_launched_without_completion_is_ng(self):
        result = check("COMPLETED", stdout="Job: [SimpleJob: [name=orderJob]] launched with the following parameters: [{}]")
        self.assertEqual(result.verdict, NG)
        self.assertIn("強制終了", result.detail)

    def test_job_filter_and_other_jobs(self):
        out = completed("prepareJob", "COMPLETED")
        result = check({"expected": "COMPLETED", "job": "orderJob"}, stdout=out)
        self.assertEqual(result.verdict, NG, "指定ジョブの完了行が無ければ、他のジョブが完了していても NG")
        self.assertIn("prepareJob", result.detail)
        both = out + "\n" + completed("orderJob", "COMPLETED")
        self.assertEqual(check({"expected": "COMPLETED", "job": "orderJob"}, stdout=both).verdict, OK)

    def test_multiple_jobs_without_filter_is_review(self):
        # Boot 2 系で spring.batch.job.names を付け忘れると全ジョブが走る。全部 COMPLETED でも自動合格にしない
        out = completed("aJob", "COMPLETED") + "\n" + completed("bJob", "COMPLETED")
        self.assertEqual(check("COMPLETED", stdout=out).verdict, REVIEW)

    def test_manual_flag_and_log_file_source(self):
        result = check({"expected": "FAILED", "manual": True}, logs=[completed("orderJob", "FAILED")])
        self.assertEqual(result.verdict, REVIEW)

    def test_console_and_file_duplicate_counted_once(self):
        line = completed("orderJob", "COMPLETED")
        result = check("COMPLETED", stdout=line, logs=[line])
        self.assertEqual(result.verdict, OK)
        self.assertEqual(result.detail, "orderJob=COMPLETED")

    def test_completed_run_does_not_hide_a_later_unfinished_run(self):
        """同じジョブを別パラメータで 2 回起動し、2 回目に完了行が無い。完了行だけを見ていた頃は OK だった。"""
        out = "\n".join([launched("orderJob", "{'targetDate':'A'}"), completed("orderJob", "COMPLETED", "'targetDate':'A'"),
                         launched("orderJob", "{'targetDate':'B'}")])
        result = check("COMPLETED", stdout=out)
        self.assertEqual(result.verdict, NG)
        self.assertIn("完了行の無い実行", result.detail)
        self.assertIn("'targetDate':'B'", result.detail)

    def test_launch_and_completion_in_console_and_file_pair_up(self):
        lines = "\n".join([launched("orderJob", "{'targetDate':'A'}"),
                           completed("orderJob", "COMPLETED", "'targetDate':'A'")])
        self.assertEqual(check("COMPLETED", stdout=lines, logs=[lines]).verdict, OK)
        # ログファイル側だけ途中で切れていても、コンソールに完了行があれば完了した実行
        self.assertEqual(check("COMPLETED", stdout=lines, logs=[launched("orderJob", "{'targetDate':'A'}")]).verdict, OK)

    def test_unfinished_run_of_another_job_needs_review(self):
        out = "\n".join([launched("prepareJob", "{}"),
                         launched("orderJob", "{'targetDate':'A'}"), completed("orderJob", "COMPLETED", "'targetDate':'A'")])
        result = check({"expected": "COMPLETED", "job": "orderJob"}, stdout=out)
        self.assertEqual(result.verdict, REVIEW)
        self.assertIn("prepareJob", result.detail)

    def test_restart_with_same_parameters_counts_each_run(self):
        out = "\n".join([launched("orderJob", "{p}"), completed("orderJob", "FAILED", "p"),
                         launched("orderJob", "{p}")])
        self.assertEqual(check("FAILED", stdout=out).verdict, NG, "2 回目の起動に完了行が無い")

    def test_schema_rejects_unknown_status_and_keys(self):
        path = Path("x.yaml")
        for spec in ("COMPLETE", {"expected": "DONE"}, {"expected": "COMPLETED", "jobb": "x"}, 1):
            problems = _validate_case_schema({"assert": {"job_status": spec}}, path, "X")
            self.assertTrue(problems, spec)
        self.assertEqual(_validate_case_schema(
            {"assert": {"job_status": {"expected": "COMPLETED", "job": "orderJob"}}}, path, "X"), [])


class JavaProject(unittest.TestCase):
    """type: java の batch を持つ最小プロジェクト。java の代わりに Python でダミーを起動する。"""

    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="autotest_java_")).resolve()
        self.jar = self.root / "app" / "order-batch.jar"
        self.jar.parent.mkdir(parents=True)
        self.jar.write_bytes(b"PK")   # 存在確認だけに使う

    def tearDown(self):
        shutil.rmtree(str(self.root), ignore_errors=True)

    def settings(self, batch, batches=None, paths=None):
        raw = {
            "env": {"sandbox": True},
            "database": {"server": "unused", "database": "unused", "auth": "windows"},
            "batch": batch,
            "paths": paths or {"input_dir": "sb/in", "output_dir": "sb/out", "log_dir": "sb/log"},
            "log": {"encoding": "utf-8"},
            "evidence": {"mode": "render"},
        }
        if batches:
            raw["batches"] = batches
        path = self.root / "config" / "settings.yaml"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(yaml.safe_dump(raw, allow_unicode=True), encoding="utf-8")
        return path, load_settings(path, self.root, require_runtime=False)


class TestJavaCommand(JavaProject):
    def test_command_order_jvm_args_jar_common_args_case_args(self):
        _path, settings = self.settings({"type": "java", "java": "/opt/jdk/bin/java", "jar": "app/order-batch.jar",
                                         "jvm_args": ["-Xmx1g", "-Dday={date}"],
                                         "common_args": ["--spring.batch.job.name=orderJob"]})
        settings.set_base_date("20260803")
        program, args = runner.build_command(settings, settings.batch, ["targetDate={date}"])
        self.assertEqual(program, Path("/opt/jdk/bin/java"))
        self.assertEqual(args, ["-Xmx1g", "-Dday=20260803", "-jar", str(self.jar),
                                "--spring.batch.job.name=orderJob", "targetDate=20260803"])

    def test_exe_command_is_unchanged(self):
        _path, settings = self.settings({"exe_path": "C:/app/Batch.exe", "common_args": ["--a"]})
        program, args = runner.build_command(settings, settings.batch, ["--b"])
        self.assertEqual(program.name, "Batch.exe")
        self.assertEqual(args, ["--a", "--b"])

    def test_java_resolution_order(self):
        batch = {"type": "java", "jar": "app/order-batch.jar"}
        with mock.patch.dict(os.environ, {"JAVA_HOME": "/env/jdk"}):
            self.assertEqual(runner._resolve_java(dict(batch, env_vars={"JAVA_HOME": "/batch/jdk"}), self.root).parent,
                             Path("/batch/jdk/bin"), "batch.env_vars の JAVA_HOME が環境変数より優先")
            self.assertEqual(runner._resolve_java(batch, self.root).parent, Path("/env/jdk/bin"))
        with mock.patch.dict(os.environ, {}, clear=True), \
                mock.patch("autotest.runner.shutil.which", return_value="/usr/bin/java"):
            self.assertEqual(runner._resolve_java(batch, self.root), Path("/usr/bin/java"))
        with mock.patch.dict(os.environ, {}, clear=True), \
                mock.patch("autotest.runner.shutil.which", return_value=None):
            self.assertIsNone(runner._resolve_java(batch, self.root))
            _path, settings = self.settings(batch)
            with self.assertRaises(ConfigError):
                runner.build_command(settings, settings.batch, [])

    def test_preflight_reports_missing_jar_and_java(self):
        _path, settings = self.settings({"type": "java", "java": str(self.root / "nojava"), "jar": "app/missing.jar"})
        problems = runner.batch_preflight(settings, settings.batch, "batch（既定）")
        self.assertTrue(any("jar が見つかりません" in p for p in problems), problems)
        self.assertTrue(any("java が見つかりません" in p for p in problems), problems)

    def test_settings_validation(self):
        bad = [
            {"type": "jar", "jar": "a.jar"},                       # 未知の type
            {"type": "java", "jar": "a.jar", "jvm_arg": ["-X"]},   # 綴り間違い
            {"type": "java", "jar": "a.jar", "jvm_args": "-Xmx1g"},  # 文字列は 1 文字ずつ分解される
        ]
        for batch in bad:
            path = self.root / "config" / "bad.yaml"
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(yaml.safe_dump({"database": {"server": "s", "database": "d"}, "batch": batch,
                                            "paths": {"input_dir": "in"}}), encoding="utf-8")
            with self.assertRaises(ConfigError, msg=batch):
                load_settings(path, self.root, require_runtime=False)
        path, _settings = self.settings({"type": "java"})
        with self.assertRaises(ConfigError) as ctx:
            load_settings(path, self.root, require_runtime=True)
        self.assertIn("batch.jar", str(ctx.exception))


class TestSpringEndToEnd(JavaProject):
    """ダミーの Spring Boot batch を実際に起動し、判定まで通す。"""

    def make_case(self, assertions, fake_args=(), params=("targetDate=20260803",), amounts=("100", "200")):
        _path, settings = self.settings(
            {"type": "java", "java": "{python}", "jvm_args": [str(FAKE)], "jar": "app/order-batch.jar",
             "common_args": ["--spring.batch.job.name=orderJob", "--fake.root=sb",
                             "--logging.file.name=sb/log/app.log"] + list(fake_args),
             "working_dir": ".", "console_encoding": "utf-8", "timeout_sec": 30, "log_settle_sec": 0},
            paths={"input_dir": "sb/in", "processed_dir": "sb/processed", "error_dir": "sb/error",
                   "output_dir": "sb/out", "log_dir": "sb/log"})
        case_dir = self.root / "cases" / "SB" / "input"
        case_dir.mkdir(parents=True, exist_ok=True)
        (case_dir / "ORDER_1.csv").write_text(
            "ORDER_NO,CUSTOMER,AMOUNT\n" + "".join("X%d,c,%s\n" % (i, a) for i, a in enumerate(amounts)),
            encoding="utf-8")
        (self.root / "cases" / "SB.yaml").write_text(yaml.safe_dump({
            "id": "SB", "name": "spring",
            "setup": {"clean_dirs": ["input_dir", "processed_dir", "error_dir", "output_dir"],
                      "input_files": [{"src": "input/ORDER_1.csv", "dest_dir": "input_dir"}]},
            "execute": {"args": list(params)},
            "collect": {"folder_evidence": []},
            "assert": assertions,
        }, allow_unicode=True), encoding="utf-8")
        case = load_cases(self.root / "cases")[0]
        self.assertEqual(preflight_case(settings, case), [])
        return settings, case

    def run_case(self, settings, case):
        return CaseRunner(settings, self.root / "output" / "run", offline=True).run(case)

    def test_completed_job_passes(self):
        settings, case = self.make_case({"exit_code": 0, "job_status": {"expected": "COMPLETED", "job": "orderJob"}})
        result = self.run_case(settings, case)
        self.assertEqual(result.verdict, OK, [(c.name, c.verdict, c.detail) for c in result.checks])
        self.assertIn("-jar %s" % self.jar, result.execution.command)

    def test_exit_code_zero_on_failure_is_caught_by_job_status(self):
        # 終了コードのマッピングが無い作り: ジョブは FAILED なのに 0 で終わる
        settings, case = self.make_case({"exit_code": 0, "job_status": "COMPLETED"},
                                        fake_args=["--fake.exit-code=always-zero"], amounts=("100", "0"))
        result = self.run_case(settings, case)
        verdicts = {c.category: c.verdict for c in result.checks}
        self.assertEqual(verdicts["exit_code"], OK, "終了コードだけ見ると合格に見える")
        self.assertEqual(verdicts["job_status"], NG)
        self.assertEqual(result.verdict, NG, "ケース全体は偽 OK にならない")

    def test_already_completed_instance_is_ng(self):
        repo = str(self.root / "sb" / "repo.json")
        settings, case = self.make_case({"job_status": "COMPLETED"},
                                        fake_args=["--fake.job-repository=" + repo, "--fake.exit-code=always-zero"])
        self.assertEqual(self.run_case(settings, case).verdict, OK)
        second = self.run_case(settings, case)   # 同じ識別パラメータで再実行 → ジョブは起動しない
        job = [c for c in second.checks if c.category == "job_status"][0]
        self.assertEqual(job.verdict, NG)
        self.assertIn("A job instance already exists", job.detail)


class TestValidateWarnsExitCodeOnly(JavaProject):
    def test_java_case_with_exit_code_only_gets_recommendation(self):
        path, _settings = self.settings({"type": "java", "java": "{python}", "jar": "app/order-batch.jar"})
        (self.root / "cases").mkdir()
        (self.root / "cases" / "J1.yaml").write_text(
            "id: J1\nname: j\nassert: {exit_code: 0}\n", encoding="utf-8")
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = cli.main(["validate", "--config", str(path)])
        self.assertEqual(code, 0, out.getvalue() + err.getvalue())
        self.assertIn("assert.job_status の併用を推奨", out.getvalue())
        self.assertIn("batch.jar", out.getvalue())


if __name__ == "__main__":
    unittest.main()
