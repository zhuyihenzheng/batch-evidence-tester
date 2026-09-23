"""Folder configuration isolation across discovery, validation and execution."""

import io
import json
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path
from unittest import mock

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import cli
from autotest.config import ConfigError, default_config, find_case_files, load_cases, load_settings
from autotest.orchestrator import CaseRunner, preflight_case


class FolderSettingsTest(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="autotest_folders_")).resolve()
        self.config = self.write("config/settings.yaml", {
            "database": {"server": "unused", "database": "unused", "auth": "windows"},
            "env": {"sandbox": True},
            "batch": {"console_encoding": "utf-8", "timeout_sec": 10,
                      "log_settle_sec": 0, "env_vars": {"COMMON": "yes"}},
            "log": {"encoding": "utf-8"},
            "evidence": {"mode": "render"},
        })
        for group in ("order", "invoice"):
            self.write("cases/%s/settings.yaml" % group, {
                "batch": {"exe_path": "{python}", "working_dir": ".",
                          "common_args": ["-c", "print('%s')" % group]},
                "paths": {"input_dir": "sandbox_%s/in" % group,
                          "output_dir": "sandbox_%s/out" % group,
                          "log_dir": "sandbox_%s/log" % group},
            })
            self.write("cases/%s/%s_001.yaml" % (group, group), {
                "id": group + "_001", "name": group, "assert": {"exit_code": 0}})
        self.settings = load_settings(self.config, self.root, require_runtime=False)

    def tearDown(self):
        shutil.rmtree(str(self.root), ignore_errors=True)

    def write(self, rel, data):
        path = self.root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
        return path

    def case(self, group="order"):
        return load_cases(self.root / "cases", only=[group + "_001"])[0]

    def test_discovery_excludes_all_config_names_and_material_yaml(self):
        self.write("cases/order/setting_local.yml", {"batch": {"timeout_sec": 20}})
        self.write("cases/order/order_001/input/data.yaml", {"arbitrary": "material"})
        self.assertEqual(len(find_case_files(self.root / "cases")), 2)
        self.assertEqual(len(load_cases(self.root / "cases")), 2)

    def test_inheritance_lists_replace_and_siblings_are_isolated(self):
        self.write("cases/settings.yaml", {"batch": {"env_vars": {"PARENT": "yes"}}})
        self.write("cases/order/settings.local.yaml", {
            "batch": {"common_args": ["-V"], "env_vars": {"ORDER": "yes"}},
            "paths": {"output_dir": "sandbox_order/local_out"}})
        order = self.settings.for_case(self.case())
        invoice = self.settings.for_case(self.case("invoice"))
        self.assertEqual(order.batch["common_args"], ["-V"])
        self.assertEqual(order.batch["env_vars"], {"COMMON": "yes", "PARENT": "yes", "ORDER": "yes"})
        self.assertEqual(order.resolve_dir("output_dir"), self.root / "sandbox_order/local_out")
        self.assertEqual(invoice.resolve_dir("output_dir"), self.root / "sandbox_invoice/out")
        self.assertNotIn("ORDER", invoice.batch["env_vars"])
        self.assertNotIn("paths", self.settings.raw)
        order.batch["env_vars"]["COMMON"] = "changed"
        self.assertEqual(self.settings.batch["env_vars"]["COMMON"], "yes")

    def test_nested_folder_and_direct_cases_dir_preserve_parent_settings(self):
        self.write("cases/order/nested/settings.yml", {"batch": {"timeout_sec": 2}})
        self.write("cases/order/nested/nested_001.yml", {"id": "nested_001"})
        case = load_cases(self.root / "cases/order/nested")[0]
        effective = self.settings.for_case(case)
        self.assertEqual(effective.batch["timeout_sec"], 2)
        self.assertEqual(effective.batch["exe_path"], "{python}")
        self.assertEqual(effective.resolve_dir("output_dir"), self.root / "sandbox_order/out")

    def test_external_cases_root_is_supported(self):
        self.write("external/settings.yaml", {"batch": {"exe_path": "{python}"}, "paths": {"log_dir": "logs"}})
        self.write("external/group/EXT.yaml", {"id": "EXT"})
        case = load_cases(self.root / "external")[0]
        self.assertEqual(self.settings.for_case(case).resolve_dir("log_dir"), self.root / "logs")

    def test_ambiguous_local_names_are_rejected(self):
        self.write("cases/order/settings.local.yaml", {})
        self.write("cases/order/setting_local.yml", {})
        with self.assertRaises(ConfigError):
            self.settings.for_case(self.case())

    def test_local_name_alias_matches_cli(self):
        local = self.write("config/setting_local.yml", self.settings.raw)
        self.assertEqual(default_config(self.root), local)
        with mock.patch.object(cli, "PROJECT_ROOT", self.root):
            self.assertEqual(cli._default_config(), local)

    def test_common_only_config_validates_after_merging(self):
        with self.assertRaises(ConfigError):
            load_settings(self.config, self.root)
        self.assertEqual(preflight_case(self.settings, self.case()), [])

    def test_invalid_folder_section_and_yaml_fail_before_setup(self):
        path = self.write("cases/order/settings.local.yaml", {"database": {"server": "other"}})
        runner = CaseRunner(self.settings, self.root / "run", offline=True)
        with mock.patch.object(runner, "_setup") as setup:
            result = runner.run(self.case())
        self.assertIn("database", result.fatal_error)
        setup.assert_not_called()
        path.write_text("batch: [\n", encoding="utf-8")
        self.assertIn(str(path), preflight_case(self.settings, self.case())[0])

    def test_invalid_section_type_has_config_error(self):
        self.write("cases/order/settings.local.yaml", {"batch": "bad"})
        with self.assertRaises(ConfigError):
            self.settings.for_case(self.case())

    def test_named_and_setup_batches_inherit_folder_settings(self):
        self.write("cases/order/settings.local.yaml", {"batches": {
            "prepare": {"common_args": ["-c", "print('prepare')"]},
            "main": {"common_args": ["-c", "print('main')"]}}})
        case = self.case()
        case.setup = {"batches": ["prepare"]}
        case.execute = {"batch": "main"}
        result = CaseRunner(self.settings, self.root / "run", offline=True).run(case)
        self.assertEqual(result.fatal_error, "")
        self.assertEqual(result.execution.stdout.strip(), "main")

    def test_reused_runner_runs_correct_executable_arguments_and_paths(self):
        runner = CaseRunner(self.settings, self.root / "run", offline=True)
        for group in ("order", "invoice", "order"):
            result = runner.run(self.case(group))
            self.assertFalse(result.fatal_error, result.fatal_error)
            self.assertEqual(result.execution.stdout.strip(), group)
            self.assertEqual(runner.settings.resolve_dir("log_dir"), self.root / ("sandbox_%s/log" % group))

    def test_distinct_exe_paths_are_selected_in_dry_run(self):
        runner = CaseRunner(self.settings, self.root / "run", dry_run=True)
        for group in ("order", "invoice"):
            exe = self.root / (group + ".exe")
            exe.write_text("placeholder", encoding="utf-8")
            self.write("cases/%s/settings.local.yaml" % group,
                       {"batch": {"exe_path": str(exe), "common_args": []}, "batches": None})
            result = runner.run(self.case(group))
            self.assertFalse(result.fatal_error, result.fatal_error)
            self.assertIn(str(exe), result.execution.command)

    def test_manual_before_after_use_same_folder_and_session_date(self):
        case = self.case()
        case.mode = "manual"
        session_dir = self.root / "manual"
        runner = CaseRunner(self.settings, session_dir, offline=True, default_date="20200101")
        session = runner.run_manual_before(case, session_dir)
        self.assertEqual(Path(session.log_dir), self.root / "sandbox_order/log")
        after_runner = CaseRunner(self.settings, session_dir, offline=True)
        result = after_runner.run_manual_after(case, session, session_dir)
        self.assertFalse(result.fatal_error, result.fatal_error)
        self.assertEqual(after_runner.settings.base_date.strftime("%Y%m%d"), "20200101")
        self.assertEqual(after_runner.settings.resolve_dir("log_dir"), Path(session.log_dir))

    def test_cli_validate_and_run_use_folder_settings_and_make_single_report(self):
        output = io.StringIO()
        with redirect_stdout(output), redirect_stderr(output):
            valid = cli.main(["validate", "--config", str(self.config)])
            ran = cli.main(["run", "--config", str(self.config), "--offline"])
        self.assertEqual(valid, 0, output.getvalue())
        self.assertEqual(ran, 0, output.getvalue())
        self.assertEqual(len(list((self.root / "output").glob("*.xlsx"))), 1)
        results = list((self.root / "output").glob("*/results/*_001.json"))
        self.assertEqual(len(results), 2)
        for path in results:
            data = json.loads(path.read_text(encoding="utf-8"))
            self.assertIn(data["case_id"].split("_")[0], data["execution"]["stdout"])

    def test_cli_folder_selection_does_not_require_other_batch_config(self):
        self.write("cases/invoice/settings.local.yaml", {"batch": "invalid"})
        output = io.StringIO()
        with redirect_stdout(output), redirect_stderr(output):
            status = cli.main(["validate", "--config", str(self.config), "--tag", "order"])
        self.assertEqual(status, 0, output.getvalue())

    def test_folder_capture_case_over_batch_over_folder_over_common(self):
        self.settings.raw["folder_evidence"] = {"targets": ["input_dir"], "max_entries": 40,
                                                 "exclude_patterns": ["*.tmp"]}
        self.write("cases/order/settings.local.yaml", {
            "folder_evidence": {"targets": ["output_dir"], "recursive": True},
            "batches": {"special": {"folder_evidence": {"targets": ["log_dir"], "max_entries": 8}}}})
        case = self.case()
        effective = self.settings.for_case(case)
        self.assertEqual(effective.folder_evidence_for(case)["targets"], ["output_dir"])
        case.execute = {"batch": "special"}
        self.assertEqual(effective.folder_evidence_for(case)["targets"], ["log_dir"])
        case.collect = {"folder_evidence": {"targets": ["input_dir", "output_dir"], "max_entries": 3}}
        cfg = effective.folder_evidence_for(case)
        self.assertEqual(cfg, {"targets": ["input_dir", "output_dir"], "max_entries": 3,
                               "exclude_patterns": ["*.tmp"], "recursive": True})
        self.assertEqual(self.settings.for_case(self.case("invoice")).folder_evidence["targets"], ["input_dir"])

    def test_empty_capture_list_disables_images_and_unneeded_directory_creation(self):
        self.settings.raw["folder_evidence"] = {"targets": ["input_dir"]}
        case = self.case()
        case.collect = {"folder_evidence": []}
        runner = CaseRunner(self.settings, self.root / "run", offline=True)
        result = runner.run(case)
        self.assertFalse(result.fatal_error, result.fatal_error)
        self.assertEqual(result.images, [])
        self.assertFalse((self.root / "sandbox_order/in").exists())
        self.assertNotIn("input_dir", runner._aliases_used_by(case))

    def test_capture_options_reach_directory_listing_and_renderer(self):
        from autotest import fsops

        self.settings.raw["folder_evidence"] = {"targets": ["input_dir"]}
        case = self.case()
        case.collect = {"folder_evidence": {"targets": ["output_dir"], "recursive": True,
                                            "exclude_patterns": ["*.tmp"], "max_entries": 5}}
        runner = CaseRunner(self.settings, self.root / "run", offline=True)
        runner._select_case_settings(case)
        with mock.patch.object(fsops, "list_dir", return_value=[]) as listing, \
                mock.patch.object(runner, "_folder_image", return_value=(self.root / "image.png", "")) as capture:
            images = runner._capture_folders(object(), "実行後", case)
        listing.assert_called_once_with(self.root / "sandbox_order/out", exclude_patterns=["*.tmp"], recursive=True)
        self.assertEqual(capture.call_args[0][-3:], ("output_dir", "実行後", 5))
        self.assertEqual(len(images), 1)
        self.assertIn("output_dir", images[0].title)

    def test_actual_before_after_images_use_per_case_targets_without_leaking(self):
        self.settings.raw["folder_evidence"] = {"targets": ["input_dir"]}
        runner = CaseRunner(self.settings, self.root / "run", offline=True)
        for group, targets in (("order", ["output_dir"]), ("invoice", ["input_dir"])):
            case = self.case(group)
            if group == "order":
                case.collect = {"folder_evidence": targets}
            result = runner.run(case)
            self.assertFalse(result.fatal_error, result.fatal_error)
            self.assertEqual([i.title for i in result.images], [
                "フォルダ確認: %s（実行前）" % targets[0], "フォルダ確認: %s（実行後）" % targets[0]])
            self.assertTrue(all(i.path.is_file() for i in result.images))

    def test_manual_capture_uses_case_override_in_both_phases(self):
        self.settings.raw["folder_evidence"] = {"targets": ["input_dir"]}
        case = self.case()
        case.mode = "manual"
        case.collect = {"folder_evidence": {"targets": ["output_dir"]}}
        session_dir = self.root / "manual"
        before = CaseRunner(self.settings, session_dir, offline=True)
        session = before.run_manual_before(case, session_dir)
        after = CaseRunner(self.settings, session_dir, offline=True)
        result = after.run_manual_after(case, session, session_dir)
        self.assertFalse(result.fatal_error, result.fatal_error)
        self.assertEqual(len(result.images), 2)
        self.assertTrue(all("output_dir" in i.title and i.path.is_file() for i in result.images))

    def test_invalid_capture_options_are_rejected_when_loading_case(self):
        for value in ("input_dir", None, [1], {"targets": "input_dir"}, {"recursive": "false"},
                      {"max_entries": 0}, {"max_entries": True}, {"target": ["input_dir"]}):
            with self.subTest(value=value):
                self.write("cases/order/order_001.yaml", {"id": "order_001", "collect": {"folder_evidence": value}})
                with self.assertRaises(ConfigError):
                    self.case()

    def test_unknown_capture_alias_fails_before_any_setup(self):
        self.write("cases/order/settings.local.yaml", {"batches": {"special": {
            "folder_evidence": {"targets": ["misspelled_dir"]}}}})
        case = self.case()
        case.execute = {"batch": "special"}
        runner = CaseRunner(self.settings, self.root / "run", offline=True)
        with mock.patch.object(runner, "_setup") as setup:
            result = runner.run(case)
        self.assertIn("misspelled_dir", result.fatal_error)
        setup.assert_not_called()


if __name__ == "__main__":
    unittest.main()
