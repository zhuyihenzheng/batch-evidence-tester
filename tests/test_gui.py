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
from autotest.layout_txt import LayoutField  # noqa: E402
from autotest import layout_inspect_gui  # noqa: E402


class ReceivedTxtGuiCase(unittest.TestCase):
    def test_failed_new_read_cannot_export_previous_results(self):
        holder = types.SimpleNamespace(
            busy=False, result=object(), saved_path=Path("old.xlsx"), window=mock.Mock(),
            _clear=mock.Mock(), _start=mock.Mock(), encoding=mock.Mock(),
            _sync_definition=mock.Mock(return_value=True),
            excel=Path("definition.xlsx"), options={})
        holder.encoding.get.return_value = "cp932"
        with mock.patch.object(layout_inspect_gui.filedialog, "askopenfilenames",
                               return_value=["broken.txt"], create=True):
            layout_inspect_gui.InspectionWindow._choose(holder)
        self.assertIsNone(holder.result)
        self.assertIsNone(holder.saved_path)
        holder._clear.assert_called_once_with()
        self.assertEqual(holder._start.call_args[0][0], "read")

    def test_export_uses_entire_inspection_not_screen_filter(self):
        result = object()
        holder = types.SimpleNamespace(
            busy=False, result=result, window=mock.Mock(), output_dir="output",
            excel=Path("definition.xlsx"), saved_path=None, _start=mock.Mock(),
            optional_detail_columns={0, 1, 2, 5, 7, 14})
        with mock.patch.object(layout_inspect_gui.filedialog, "asksaveasfilename",
                               return_value="check.xlsx", create=True), \
                mock.patch.object(layout_inspect_gui, "export_inspection") as export:
            layout_inspect_gui.InspectionWindow._export(holder)
            holder._start.call_args[0][1]()
        export.assert_called_once_with(result, "check.xlsx", overwrite=True,
                                       detail_columns=(0, 1, 2, 5, 6, 7, 8, 11, 12, 14))

    def test_export_passes_selected_optional_columns(self):
        result = object()
        holder = types.SimpleNamespace(
            busy=False, result=result, window=mock.Mock(), output_dir="output",
            excel=Path("definition.xlsx"), saved_path=None, _start=mock.Mock(),
            optional_detail_columns={7, 10, 14})
        with mock.patch.object(layout_inspect_gui.filedialog, "asksaveasfilename",
                               return_value="check.xlsx", create=True), \
                mock.patch.object(layout_inspect_gui, "export_inspection") as export:
            layout_inspect_gui.InspectionWindow._export(holder)
            holder._start.call_args[0][1]()
        export.assert_called_once_with(result, "check.xlsx", overwrite=True,
                                       detail_columns=(6, 7, 8, 10, 11, 12, 14))

    def test_selected_columns_update_detail_view_and_export_order(self):
        holder = types.SimpleNamespace(optional_detail_columns={7, 20, 21},
                                       detail_tree=mock.Mock(), _render=mock.Mock())
        layout_inspect_gui.InspectionWindow._refresh_detail_columns(holder)
        self.assertEqual(holder.detail_indexes, (6, 7, 8, 11, 12, 20, 21))
        holder.detail_tree.configure.assert_called_once_with(
            columns=["0", "1", "2", "3", "4", "5", "6"])
        self.assertEqual(holder.detail_tree.heading.call_args_list[5].args,
                         ("5",))
        self.assertEqual(holder.detail_tree.heading.call_args_list[5].kwargs,
                         {"text": "データ型"})
        holder._render.assert_called_once_with()
        self.assertEqual(layout_inspect_gui.selected_detail_columns({7, 20, 21}),
                         holder.detail_indexes)

    def test_detail_display_limit_is_explicit_and_does_not_change_result(self):
        row = [""] * 29
        row[0], row[1], row[13] = "received.txt", 1, "注意点あり"
        row[14] = "異常"
        record = dict(file="received.txt", number=1, form="1001", presence="1", status="注意点あり", issues="異常")
        result = types.SimpleNamespace(details=[row] * 3001, records=[record], encoding="cp932")
        holder = types.SimpleNamespace(
            busy=False, result=result, _clear=mock.Mock(), query=mock.Mock(),
            issues_only=mock.Mock(), detail_tree=mock.Mock(), record_tree=mock.Mock(),
            detail_indexes=[0, 1, 13], status=mock.Mock())
        holder.query.get.return_value = ""
        holder.issues_only.get.return_value = True
        layout_inspect_gui.InspectionWindow._render(holder)
        self.assertEqual(holder.detail_tree.insert.call_count, 3000)
        self.assertEqual(len(result.details), 3001)
        self.assertIn("明細3000/3001", holder.status.set.call_args[0][0])


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
    def test_append_current_form_keeps_selected_image_and_screen_edits(self):
        item = layout_txt_gui.PackageItem(
            base_name="shared_", form_id="1001", front_image_bytes=b"original image",
            front_recognition_text='"1001","1","1","A","0","0,0,0,1"\r\n')
        field = LayoutField("2001", "1", "1", "Item", "文字列", "", 20, "B", 3)
        holder = types.SimpleNamespace(
            _finish_package_cell_edit=mock.Mock(), package_tree=mock.Mock(),
            form_fields={"2001": [field]}, package_items={"image": item},
            form_var=mock.Mock(), loaded_fields=[field],
            _screen_edits=mock.Mock(return_value=([3], {3: {"value": "編集値"}})),
            status_var=mock.Mock(), root=object(),
            _package_source=layout_txt_gui.LayoutTxtGui._package_source)
        holder.form_var.get.return_value = "2001"
        holder.package_tree.selection.return_value = ["image"]
        holder.package_tree.set.return_value = "1"
        holder._package_values = lambda value, include: layout_txt_gui.LayoutTxtGui._package_values(
            holder, value, include)
        layout_txt_gui.LayoutTxtGui._append_current_form_to_image(holder)
        self.assertEqual(item.front_form_count, 2)
        self.assertIn('"2001","1","1","編集値"', item.front_recognition_text)
        self.assertEqual(item.front_image_bytes, b"original image")
        self.assertEqual(item.form_id, "1001")
        displayed = dict(zip(layout_txt_gui.PACKAGE_COLUMNS, holder._package_values(item, "1")))
        self.assertEqual(displayed["front_form_count"], 2)
        self.assertEqual(displayed["front_recognition"], "shared_F.txt")

    def test_csv_defaults_and_random_template_survive_settings_reload(self):
        expected = {
            "package_scan_batch_id_var": "0123456789001",
            "package_arrival_date_var": "20260907",
            "package_form_id_var": "1001",
            "package_application_number_var": "0001",
            "package_reception_number_var": "0002",
            "package_format_id_var": "01",
            "package_delivery_date_var": "20260908",
            "package_delivery_shot_var": "02",
            "filename_template_var": "{random9}0001",
        }
        variables = {name: mock.Mock() for name in layout_txt_gui.SETTING_VARIABLE_NAMES}
        for name, variable in variables.items():
            variable.get.return_value = expected.get(name, "")
        holder = types.SimpleNamespace(visible_columns=["item_name"], **variables)
        with tempfile.TemporaryDirectory() as directory:
            holder.settings_path = Path(directory) / "layout_txt_gui.json"
            layout_txt_gui.save_settings(
                layout_txt_gui.LayoutTxtGui._settings_payload(holder), holder.settings_path)
            layout_txt_gui.LayoutTxtGui._load_persisted_settings(holder)
        for name, value in expected.items():
            variables[name].set.assert_called_once_with(value)

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


@unittest.skipUnless(os.name == "nt", "Actual Windows Tk layout check")
class LayoutSmallScreenCase(unittest.TestCase):
    def test_controls_fit_without_page_horizontal_scrolling(self):
        with tempfile.TemporaryDirectory() as directory:
            def descendants(widget):
                for child in widget.winfo_children():
                    yield child
                    for nested in descendants(child):
                        yield nested

            # 1280px screen with window margins, plus resizing and larger fonts.
            for scaling, width in ((1.333, 1240), (1.333, 1000), (1.667, 1240)):
                with self.subTest(scaling=scaling, width=width):
                    root = layout_txt_gui.tk.Tk()
                    self.addCleanup(root.destroy)
                    root.tk.call("tk", "scaling", scaling)
                    with mock.patch.object(
                            layout_txt_gui, "default_settings_path",
                            return_value=Path(directory) / "settings.json"):
                        app = layout_txt_gui.LayoutTxtGui(root)
                    root.report_callback_exception = mock.Mock()
                    root.geometry("%dx668+0+0" % width)
                    app.content_canvas.yview_moveto(0.0)
                    root.update()
                    canvas = app.content_canvas
                    self.assertEqual(canvas.xview(), (0.0, 1.0))
                    left = canvas.winfo_rootx()
                    right = left + canvas.winfo_width()
                    for widget in descendants(app.content_frame):
                        if widget.winfo_class() not in (
                                "TButton", "TEntry", "TCombobox", "TCheckbutton"):
                            continue
                        label = "%s %s" % (widget, widget.winfo_class())
                        self.assertGreater(widget.winfo_width(), 1, label)
                        self.assertGreaterEqual(widget.winfo_rootx(), left, label)
                        self.assertLessEqual(
                            widget.winfo_rootx() + widget.winfo_width(), right, label)
                    for button in (app.all_button, app.selected_button):
                        self.assertLessEqual(
                            button.winfo_rooty() + button.winfo_height(),
                            root.winfo_rooty() + root.winfo_height())
                    self.assertLess(canvas.yview()[1], 1.0)
                    canvas.yview_moveto(1.0)
                    root.update()
                    self.assertAlmostEqual(canvas.yview()[1], 1.0, places=2)
                    root.report_callback_exception.assert_not_called()
                    self.doCleanups()


if __name__ == "__main__":
    unittest.main()
