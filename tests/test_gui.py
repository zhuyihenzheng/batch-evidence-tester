# -*- coding: utf-8 -*-
"""操作画面の子プロセス出力に関する回帰テスト。"""

import os
import queue
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

# CI や macOS の最小 Python には tkinter が無いことがある。ここで検査するのは
# GUI 部品ではなく subprocess の UTF-8 パイプなので、import 用の代用品で十分。
try:
    import tkinter  # noqa: F401
except ImportError:
    tkinter_stub = types.ModuleType("tkinter")
    tkinter_stub.filedialog = types.SimpleNamespace()
    tkinter_stub.messagebox = types.SimpleNamespace()
    tkinter_stub.ttk = types.SimpleNamespace()
    sys.modules["tkinter"] = tkinter_stub

from autotest import gui  # noqa: E402
from autotest import layout_txt_gui  # noqa: E402
from autotest.layout_tar import next_numbered_base_name  # noqa: E402


class GuiSubprocessEncodingCase(unittest.TestCase):
    def test_utf8_output_does_not_use_windows_locale(self):
        """UTF-8 の日本語を cp932 locale に依存せず読み取れること。"""
        text = "設定確認：接続成功"
        proc = gui._popen(
            [sys.executable, "-c", "print(%r)" % text], ROOT)
        holder = types.SimpleNamespace(queue=queue.Queue())

        gui.AutotestGui._reader(holder, proc)

        self.assertEqual(holder.queue.get(), ("line", text + os.linesep))
        self.assertEqual(holder.queue.get(), ("exit", 0))


class GuiCaseSelectionCase(unittest.TestCase):
    def test_refresh_cases_leaves_every_case_unchecked(self):
        """初回表示・再読込ではケースを自動選択しないこと。"""
        cases = [
            types.SimpleNamespace(case_id="TC001", tags=["normal"]),
            types.SimpleNamespace(case_id="TC002", tags=["error"]),
        ]
        settings = types.SimpleNamespace(database={})
        holder = types.SimpleNamespace(
            config_path=ROOT / "config" / "settings.yaml",
            project_root=ROOT,
            cases_dir=ROOT / "cases",
            settings=None,
            config_label=mock.Mock(),
            tag_box={"values": []},
            tag_var=mock.Mock(),
            checked={"TC001", "TC002"},
            db_state="unknown",
            db_actual="",
        )
        holder.tag_var.get.return_value = "（すべて）"
        holder._db_signature = gui.AutotestGui._db_signature
        holder._render_db_bar = mock.Mock()
        holder._render_rows = mock.Mock()

        with mock.patch.object(gui, "load_settings", return_value=settings), \
                mock.patch.object(gui, "load_cases", return_value=cases):
            gui.AutotestGui.refresh_cases(holder)

        self.assertEqual(holder.checked, set())
        holder._render_rows.assert_called_once_with()


class LayoutPackageFilenameCase(unittest.TestCase):
    def test_current_form_base_name_uses_three_digit_sequence(self):
        """同じ基礎名のFORMを追加すると _001_、_002_、_003_ と採番する。"""
        used = []
        names = []
        for _sequence in range(3):
            base = next_numbered_base_name("BASE_NAME_", used)
            names.append(base + "F.TIF")
            used.append(base)

        self.assertEqual(names, [
            "BASE_NAME_001_F.TIF",
            "BASE_NAME_002_F.TIF",
            "BASE_NAME_003_F.TIF",
        ])

    def test_existing_package_requires_confirmation_before_overwrite(self):
        with tempfile.TemporaryDirectory(prefix="package_confirm_") as temp_dir:
            output = Path(temp_dir)
            (output / "renamed.tar").write_bytes(b"old")
            holder = types.SimpleNamespace(
                overwrite_var=mock.Mock(), root=object(), status_var=mock.Mock())
            holder.overwrite_var.get.return_value = True

            with mock.patch.object(
                    layout_txt_gui.messagebox, "askyesno",
                    return_value=False, create=True) as confirm:
                accepted = layout_txt_gui.LayoutTxtGui._confirm_package_overwrite(
                    holder, output, "renamed")

            self.assertFalse(accepted)
            confirm.assert_called_once()
            self.assertIn("renamed.tar", confirm.call_args.args[1])
            holder.status_var.set.assert_called_once_with(
                "TARの上書きをキャンセルしました。")

    def test_new_package_does_not_show_overwrite_confirmation(self):
        with tempfile.TemporaryDirectory(prefix="package_confirm_") as temp_dir:
            holder = types.SimpleNamespace(
                overwrite_var=mock.Mock(), root=object(), status_var=mock.Mock())
            holder.overwrite_var.get.return_value = True

            with mock.patch.object(
                    layout_txt_gui.messagebox, "askyesno",
                    create=True) as confirm:
                accepted = layout_txt_gui.LayoutTxtGui._confirm_package_overwrite(
                    holder, Path(temp_dir), "new_package")

            self.assertTrue(accepted)
            confirm.assert_not_called()


if __name__ == "__main__":
    unittest.main()
