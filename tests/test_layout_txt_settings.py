# -*- coding: utf-8 -*-

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest.layout_txt_settings import (  # noqa: E402
    CONFIG_ENV,
    LayoutGuiSettingsError,
    default_settings_path,
    load_settings,
    save_settings,
)


class TestLayoutTxtSettings(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="layout_gui_settings_"))
        self.path = self.tmp / "nested" / "layout_txt_gui.json"

    def tearDown(self):
        shutil.rmtree(str(self.tmp), ignore_errors=True)

    def test_round_trip_preserves_unicode_and_boolean_values(self):
        expected = {
            "version": 1,
            "values": {"filename_template_var": "帳票_{form_id}", "create_tar_var": True},
            "visible_columns": ["item_name", "notes"],
        }
        self.assertEqual(save_settings(expected, self.path), self.path)
        self.assertEqual(load_settings(self.path), expected)
        raw = self.path.read_text(encoding="utf-8")
        self.assertIn("帳票_{form_id}", raw)
        self.assertTrue(raw.endswith("\n"))

    def test_environment_can_override_default_path(self):
        with mock.patch.dict("os.environ", {CONFIG_ENV: str(self.path)}):
            self.assertEqual(default_settings_path(), self.path)

    def test_invalid_json_is_reported(self):
        self.path.parent.mkdir(parents=True)
        self.path.write_text("{invalid", encoding="utf-8")
        with self.assertRaises(LayoutGuiSettingsError):
            load_settings(self.path)

    def test_saved_file_is_valid_json_object(self):
        save_settings({"version": 1}, self.path)
        with self.path.open("r", encoding="utf-8") as stream:
            self.assertEqual(json.load(stream), {"version": 1})


if __name__ == "__main__":
    unittest.main()
