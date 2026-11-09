# -*- coding: utf-8 -*-
"""セル範囲編集で値の欠落や部分適用が起きないことを確認する。"""

import sys
import csv
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'src'))
from autotest.layout_grid import clipboard_text, parse_clipboard, paste_changes

try:
    import tkinter as tk
except ImportError:
    tk = None


class ClipboardTests(unittest.TestCase):
    def test_excel_roundtrip_retains_whitespace_zeros_quotes_tabs_newlines(self):
        values = [['001', ' 文字 ', '"引用"', 'a\tb', '行1\n行2', ''], ['', '0', '', '', '', '']]
        self.assertEqual(parse_clipboard(clipboard_text(values)), values)

    def test_blank_and_trailing_cells(self):
        self.assertEqual(parse_clipboard(''), [['']])
        self.assertEqual(parse_clipboard('a\t\r\n\t\r\n'), [['a', ''], ['', '']])
        with self.assertRaises(ValueError):
            parse_clipboard('a\tb\nc')

    def test_paste_from_single_cell_and_fill_rectangle(self):
        changes, bounds = paste_changes(['r1', 'r2'], ['value', 'attr'], (0, 0, 0, 0),
                                        [['001', '0'], ['002', '2']], {'value', 'attr'})
        self.assertEqual(changes[('r2', 'value')], '002')
        self.assertEqual(bounds, (0, 0, 1, 1))
        changes, _ = paste_changes(['r1', 'r2'], ['value', 'attr'], bounds, [['']], {'value', 'attr'})
        self.assertEqual(set(changes.values()), {''})

    def test_readonly_overflow_and_mismatched_selection_are_rejected(self):
        for columns, bounds, values in (
                (['readonly', 'value'], (0, 0, 0, 0), [['a', 'b']]),
                (['value'], (0, 0, 0, 0), [['a', 'b']]),
                (['value'], (1, 0, 1, 0), [['a'], ['b']]),
                (['value'], (0, 0, 1, 0), [['a'], ['b'], ['c']])):
            with self.subTest(columns=columns, bounds=bounds), self.assertRaises(ValueError):
                paste_changes(['r1', 'r2'], columns, bounds, values, {'value'})


@unittest.skipUnless(tk is not None and hasattr(tk, 'Tk'), 'Tk is required')
class GridGuiTests(unittest.TestCase):
    def setUp(self):
        from autotest.layout_txt_gui import LayoutTxtGui
        from autotest.layout_txt import LayoutField
        self.temp = tempfile.TemporaryDirectory()
        self.root = tk.Tk()
        self.root.withdraw()
        with mock.patch('autotest.layout_txt_gui.default_settings_path',
                        return_value=Path(self.temp.name) / 'settings.json'):
            self.app = LayoutTxtGui(self.root)
        self.fields = [LayoutField('1001', '1', str(i), '項目', '文字列', '', 10,
                                   str(i), i, '0', '10,10,50,50') for i in (1, 2, 3)]
        self.app.loaded_fields = self.fields
        self.app.form_fields = {'1001': self.fields}
        self.app.form_var.set('1001')
        self.app._show_selected_form()
        self.grid = self.app.field_grid

    def tearDown(self):
        self.root.update()
        self.root.destroy()
        self.temp.cleanup()

    def paste(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.grid.paste()

    def test_bulk_paste_copy_clear_undo_redo_and_hidden_columns(self):
        self.app.tree.configure(displaycolumns=('value', 'attribute_flag', 'coordinates'))
        self.grid.select(('1:1', 'value'))
        self.paste('001\t2\t10,10,50,50\r\n002\t0\t70,10,50,50\r\n')
        self.assertEqual(self.app.tree.set('2:1', 'value'), '002')
        self.grid.copy()
        self.assertEqual(parse_clipboard(self.root.clipboard_get())[1], ['002', '0', '70,10,50,50'])
        self.grid.select(('1:1', 'value'))
        self.grid.select(('2:1', 'value'), True)
        self.grid.clear()
        for iid in ('1:1', '2:1'):
            self.assertEqual([self.app.tree.set(iid, c) for c in ('value', 'attribute_flag', 'coordinates')], ['', '', ''])
        self.grid.undo()
        self.assertEqual(self.app.tree.set('2:1', 'coordinates'), '70,10,50,50')
        self.grid.redo()
        self.assertEqual(self.app.tree.set('2:1', 'value'), '')
        self.grid.undo()
        self.grid.undo()
        self.assertEqual(self.app.tree.set('2:1', 'value'), '2')

    def test_invalid_paste_is_atomic_and_does_not_enter_history(self):
        self.app.tree.configure(displaycolumns=('include', 'value'))
        self.grid.select(('1:1', 'include'))
        with mock.patch('tkinter.messagebox.showerror') as error:
            self.paste('0\t変更\ninvalid\t変更')
        error.assert_called_once()
        self.assertEqual(self.app.tree.set('1:1', 'include'), '1')
        self.assertEqual(self.app.tree.set('1:1', 'value'), '1')
        self.assertEqual(self.grid.undo_stack, [])

    def test_package_paste_updates_model_and_undo_restores_derived_names(self):
        from autotest.layout_tar import PackageItem
        for name in ('a', 'b'):
            self.app._insert_package_item(PackageItem(name, front_image_bytes=b'img'))
        grid, tree = self.app.package_grid, self.app.package_tree
        tree.configure(displaycolumns=('base_name', 'scan_batch_id'))
        grid.select(('package_00001', 'base_name'))
        self.root.clipboard_clear()
        self.root.clipboard_append('new_a\t001\nnew_b\t002')
        grid.paste()
        self.assertEqual(self.app.package_items['package_00002'].scan_batch_id, '002')
        self.assertEqual(tree.set('package_00001', 'front_image'), 'new_aF.tif')
        grid.undo()
        self.assertEqual(tree.set('package_00001', 'front_image'), 'aF.tif')
        grid.redo()
        self.assertEqual(tree.set('package_00001', 'front_image'), 'new_aF.tif')
        grid.select(('package_00001', 'base_name'))
        grid.select(('package_00002', 'base_name'), True)
        with mock.patch('tkinter.messagebox.showerror') as error:
            grid.clear()
        error.assert_called_once()
        self.assertEqual(self.app.package_items['package_00001'].base_name, 'new_a')

    def test_single_cell_edit_and_tab_navigation_use_same_undo(self):
        self.app.tree.configure(displaycolumns=('value', 'attribute_flag'))
        # 非表示ウィンドウでも編集処理を検査できるよう、セル領域だけを固定する。
        with mock.patch.object(self.app.tree, 'bbox', return_value=(0, 25, 100, 20)):
            self.grid.select(('1:1', 'value'))
            self.grid.edit(initial='編集')
            self.grid.end_edit(0, 1)
        self.assertEqual(self.app.tree.set('1:1', 'value'), '編集')
        self.assertEqual(self.grid.active, ('1:1', 'attribute_flag'))
        self.grid.undo()
        self.assertEqual(self.app.tree.set('1:1', 'value'), '1')

    def test_form_switch_resets_old_history(self):
        self.grid.select(('1:1', 'value'))
        self.grid.clear()
        self.assertTrue(self.grid.undo_stack)
        self.app._show_selected_form()
        self.assertEqual(self.grid.undo_stack, [])
        self.assertEqual(self.app.tree.set('1:1', 'value'), '')

    def add_editable_package(self):
        from autotest.layout_tar import PackageItem
        stream = io.StringIO()
        csv.writer(stream, quoting=csv.QUOTE_ALL, lineterminator='\r\n').writerows([
            ['1001', '1', '1', '元の値', '0', '10,10,50,50'],
            ['1001', '1', '1', '別の帳票', '2', '70,10,50,50']])
        item = PackageItem('saved_', form_id='999', front_image_bytes=b'image',
                           front_recognition_text=stream.getvalue())
        iid = self.app._insert_package_item(item)
        self.app.package_tree.set(iid, 'include', '0')
        self.app._edit_package_content()
        editor = self.app.package_content_editor
        editor.window.withdraw()
        return iid, item, editor

    def test_package_contents_save_multiple_forms_and_persist_in_template(self):
        from autotest.layout_templates import save_template, load_template
        iid, item, editor = self.add_editable_package()
        original = item.front_recognition_text
        editor.form_var.set('2002')
        editor.presence_var.set('0')
        editor.grid.change({('0', 'value'): '001\n"手編集",値'})
        self.assertEqual(item.front_recognition_text, original)
        editor.form_box.current(1)
        editor.switch()
        self.assertEqual(editor.form_var.get(), '1001')
        editor.grid.select(('0', 'value'))
        editor.grid.clear()
        editor.form_box.current(0)
        editor.switch()
        self.assertEqual(editor.tree.set('0', 'value'), '001\n"手編集",値')
        editor.save()
        records = list(csv.reader(io.StringIO(item.front_recognition_text)))
        self.assertEqual(records[0], ['2002', '0', '1', '001\n"手編集",値', '0', '10,10,50,50'])
        self.assertEqual(records[1], ['1001', '1', '1', '', '', ''])
        self.assertEqual(item.form_id, '999')
        self.assertEqual(item.front_image_bytes, b'image')
        self.assertEqual(self.app.package_tree.set(iid, 'include'), '0')
        path = Path(self.temp.name) / 'edited.xlsx'
        save_template(path, 'edited', {}, [], [(item, '0')])
        restored = load_template(path, 'edited')[2][0][0]
        self.assertEqual(restored.front_recognition_text, item.front_recognition_text)

    def test_package_contents_cancel_does_not_change_list(self):
        _iid, item, editor = self.add_editable_package()
        original = item.front_recognition_text
        editor.grid.change({('0', 'value'): '破棄する編集'})
        editor.form_box.current(1)
        editor.switch()
        with mock.patch('tkinter.messagebox.askyesno', return_value=True):
            editor.cancel()
        self.assertEqual(item.front_recognition_text, original)

    def test_package_contents_reject_stale_editor(self):
        _iid, item, first = self.add_editable_package()
        first.grid.change({('0', 'value'): '古い編集'})
        self.app._edit_package_content()
        second = self.app.package_content_editor
        second.window.withdraw()
        second.grid.change({('0', 'value'): '新しい編集'})
        second.save()
        with mock.patch('tkinter.messagebox.showerror') as error:
            first.save()
        error.assert_called_once()
        self.assertIn('新しい編集', item.front_recognition_text)
        self.assertNotIn('古い編集', item.front_recognition_text)
        first.window.destroy()

    def test_package_contents_invalid_form_is_not_saved(self):
        _iid, item, editor = self.add_editable_package()
        original = item.front_recognition_text
        editor.form_var.set('invalid')
        with mock.patch('tkinter.messagebox.showerror') as error:
            editor.save()
        error.assert_called_once()
        self.assertEqual(item.front_recognition_text, original)
        self.assertTrue(editor.window.winfo_exists())
        editor.window.destroy()

    def test_custom_dimensions_apply_to_selected_rows_and_empty_ocr(self):
        self.app.coordinate_width_var.set('80')
        self.app.coordinate_height_var.set('30')
        self.grid.select(('1:1', 'coordinates'))
        self.app._apply_coordinate_dimensions()
        self.assertEqual(self.app.tree.set('1:1', 'coordinates'), '10,10,80,30')
        self.assertEqual(self.app.tree.set('2:1', 'coordinates'), '10,10,50,50')
        self.grid.undo()
        self.assertEqual(self.app.tree.set('1:1', 'coordinates'), '10,10,50,50')
        self.app.tree.set('1:1', 'value', '')
        self.app._apply_coordinate_dimensions()
        self.assertEqual(self.app.tree.set('1:1', 'coordinates'), '')
        _iid, item, editor = self.add_editable_package()
        self.assertEqual(editor.width_var.get(), '80')
        self.assertEqual(editor.height_var.get(), '30')
        editor.resize()
        editor.save()
        records = list(csv.reader(io.StringIO(item.front_recognition_text)))
        self.assertEqual(records[0][-1], '10,10,80,30')
        self.assertEqual(records[1][-1], '70,10,50,50')


if __name__ == '__main__':
    unittest.main()
