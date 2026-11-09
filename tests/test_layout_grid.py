# -*- coding: utf-8 -*-
"""セル範囲編集で値の欠落や部分適用が起きないことを確認する。"""

import sys
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


if __name__ == '__main__':
    unittest.main()
