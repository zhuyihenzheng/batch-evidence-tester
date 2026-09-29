# -*- coding: utf-8 -*-
"""生成アプリなしで起動・配布できることを検証する。"""

import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

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
        self.assertEqual(result.stdout.strip(), "ReceivedTxtChecker 0.1.0")

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
