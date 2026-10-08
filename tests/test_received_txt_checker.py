# -*- coding: utf-8 -*-
"""生成アプリなしで起動・配布できることを検証する。"""

import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import received_txt_checker


class StandaloneCheckerCase(unittest.TestCase):
    def test_source_bundle_runs_outside_repository(self):
        from tools.package_received_txt_checker import package
        with tempfile.TemporaryDirectory() as directory:
            archive_path = package(Path(directory) / "checker.zip")
            with zipfile.ZipFile(str(archive_path)) as archive:
                self.assertNotIn("ReceivedTxtChecker/src/autotest/layout_txt_gui.py", archive.namelist())
                archive.extractall(directory)
            env = dict(os.environ)
            env.pop("PYTHONPATH", None)
            code = ("import received_txt_checker as app; app._prepare_source_path(); "
                    "raise SystemExit(app._functional_smoke_test())")
            result = subprocess.run([sys.executable, "-c", code],
                                    cwd=str(Path(directory) / "ReceivedTxtChecker"), env=env,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
            self.assertEqual(result.returncode, 0, result.stderr)

    def test_version_works_without_importing_tkinter(self):
        result = subprocess.run([sys.executable, str(ROOT / "received_txt_checker.py"), "--version"],
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "ReceivedTxtChecker 0.1.12")

    def test_cleanup_retries_transient_windows_sharing_violation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "smoke"
            root.mkdir()
            remove = received_txt_checker.shutil.rmtree
            calls = []

            def temporarily_locked(path):
                calls.append(path)
                if len(calls) == 1:
                    raise PermissionError(32, "file is being used", "check.xlsx")
                remove(path)
            with mock.patch.object(received_txt_checker.shutil, "rmtree", temporarily_locked), \
                    mock.patch.object(received_txt_checker.time, "sleep") as sleep:
                self.assertTrue(received_txt_checker._cleanup_smoke_directory(root))
            self.assertEqual(len(calls), 2)
            sleep.assert_called_once()
            self.assertFalse(root.exists())

    def test_persistent_sharing_violation_does_not_fail_completed_smoke(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "smoke"
            root.mkdir()
            with mock.patch.object(received_txt_checker.tempfile, "mkdtemp", return_value=str(root)), \
                    mock.patch.object(received_txt_checker.shutil, "rmtree",
                                      side_effect=PermissionError(32, "file is being used")) as remove, \
                    mock.patch.object(received_txt_checker.time, "sleep"), \
                    mock.patch.object(received_txt_checker.sys, "stderr", None):
                received_txt_checker._prepare_source_path()
                self.assertEqual(received_txt_checker._functional_smoke_test(), 0)
                self.assertEqual(remove.call_count, 10)
            self.assertTrue((root / "check.xlsx").is_file())

    def test_cleanup_does_not_hide_processing_errors(self):
        received_txt_checker._prepare_source_path()
        from autotest import layout_inspect
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "smoke"
            root.mkdir()
            with mock.patch.object(received_txt_checker.tempfile, "mkdtemp", return_value=str(root)), \
                    mock.patch.object(layout_inspect, "inspect_txt", side_effect=RuntimeError("read failed")), \
                    mock.patch.object(received_txt_checker, "_cleanup_smoke_directory", return_value=False), \
                    mock.patch.object(received_txt_checker.sys, "stderr", None):
                with self.assertRaisesRegex(RuntimeError, "read failed"):
                    received_txt_checker._functional_smoke_test()

    def test_non_sharing_cleanup_errors_are_not_ignored(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "smoke"
            root.mkdir()
            with mock.patch.object(received_txt_checker.shutil, "rmtree",
                                   side_effect=PermissionError(13, "permission denied")):
                with self.assertRaises(PermissionError):
                    received_txt_checker._cleanup_smoke_directory(root)

    def test_smoke_is_independent_of_generator_gui_and_image_packages(self):
        code = ("import sys; "
                "sys.modules['autotest.layout_txt_gui'] = None; "
                "sys.modules['autotest.layout_tar'] = None; "
                "sys.modules['PIL'] = None; "
                "import received_txt_checker as app; app._prepare_source_path(); "
                "sys.exit(app._functional_smoke_test())")
        result = subprocess.run([sys.executable, "-c", code], cwd=str(ROOT),
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    unittest.main()
