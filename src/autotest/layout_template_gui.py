# -*- coding: utf-8 -*-
"""Excelテンプレートの画面操作。"""

from collections import OrderedDict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .layout_templates import list_templates, save_template, load_template, delete_template


class TemplateGuiMixin(object):
    def _build_template_controls(self, parent):
        from .layout_txt_gui import _wrap_controls
        self.template_path = None
        self.template_mode = False
        self.editor_states = {}
        self.displayed_form = None
        frame = ttk.Frame(parent)
        frame.grid(row=5, column=0, columnspan=8, sticky='we', pady=(6, 0))
        ttk.Button(frame, text='テンプレートExcel...', command=self._choose_template).pack(side='left')
        self.template_box = ttk.Combobox(frame, state='readonly', width=20)
        self.template_box.pack(side='left', padx=5)
        for text, command in (
                ('読込', self._load_template), ('新規保存', self._new_template),
                ('上書き保存', self._save_template), ('削除', self._delete_template)):
            ttk.Button(frame, text=text, command=command).pack(side='left', padx=2)
        _wrap_controls(frame)

    def _choose_template(self):
        chosen = filedialog.askopenfilename(parent=self.root, title='テンプレートExcelを選択',
                                           filetypes=(('Excelテンプレート', '*.xlsx'),))
        if not chosen:
            return
        try:
            names = list_templates(chosen)
        except Exception as exc:
            messagebox.showerror('テンプレート読込エラー', str(exc), parent=self.root)
            return
        self.template_path = Path(chosen)
        self.template_box.configure(values=names)
        self.template_box.set(names[0] if names else '')
        self.status_var.set('テンプレートExcel: %s / 名前を選んで「読込」を押してください。' % chosen)

    def _capture_editor(self):
        self._finish_cell_edit(save=True)
        if self.displayed_form:
            self.editor_states[self.displayed_form] = {
                iid: tuple(self.tree.item(iid, 'values')) for iid in self.tree.get_children('')}

    def _template_fields(self):
        self._capture_editor()
        result = []
        for field in self.loaded_fields:
            data = dict(vars(field))
            data['include'] = '1'
            values = self.editor_states.get(field.form_id, {}).get(field.instance_key)
            if values:
                for key, index in (('include', 0), ('field_id', 3), ('value', 8),
                                   ('attribute_flag', 9), ('coordinates', 10)):
                    data[key] = values[index]
            if data['value'] == '':
                data['attribute_flag'] = data['coordinates'] = ''
            result.append(data)
        return result

    def _new_template(self):
        from tkinter import simpledialog
        name = simpledialog.askstring('新規テンプレート', 'テンプレート名:', parent=self.root)
        if name:
            self._save_template(name=name.strip())

    def _save_template(self, name=None):
        from .layout_txt_gui import SETTING_VARIABLE_NAMES
        overwrite = name is None
        name = self.template_box.get() if name is None else name
        if not name:
            messagebox.showwarning('テンプレート', '「新規保存」で名前を指定してください。', parent=self.root)
            return
        if self.template_path is None:
            chosen = filedialog.asksaveasfilename(
                parent=self.root, title='テンプレートExcelの保存先',
                defaultextension='.xlsx', initialfile='出力テンプレート.xlsx',
                filetypes=(('Excelテンプレート', '*.xlsx'),))
            if not chosen:
                return
            self.template_path = Path(chosen)
        if overwrite and not messagebox.askyesno(
                'テンプレート上書き', '%s を現在の設定で上書きしますか？' % name, parent=self.root):
            return
        try:
            self._finish_package_cell_edit(save=True)
            fields = self._template_fields()
            if not fields and not self.package_items:
                raise ValueError('Excel定義を読み込むか、出力リストへ画像を追加してください。')
            settings = {key: getattr(self, key).get() for key in SETTING_VARIABLE_NAMES}
            settings['form_var'] = self.form_var.get()
            # Excel内の説明列は手編集時の案内として使う。
            descriptions = self._setting_descriptions()
            save_template(self.template_path, name, settings, fields,
                          [(self.package_items[iid], self.package_tree.set(iid, 'include'))
                           for iid in self.package_tree.get_children('')],
                          descriptions=descriptions, overwrite=overwrite)
            self.template_box.configure(values=list_templates(self.template_path))
            self.template_box.set(name)
            self.status_var.set('テンプレート保存完了: %s / %s' % (self.template_path, name))
            messagebox.showinfo('テンプレート保存完了',
                                'Excelの各セルを編集できます。\n画像資材フォルダも一緒に保管してください。\n\n%s'
                                % self.template_path, parent=self.root)
        except Exception as exc:
            messagebox.showerror('テンプレート保存エラー', str(exc), parent=self.root)

    def _setting_descriptions(self):
        descriptions = {
            'excel_var': '元の帳票定義Excel', 'output_var': '出力フォルダ',
            'sheet_var': '元の帳票定義Sheet', 'header_var': '見出し行',
            'form_var': '表示中FORM_ID', 'use_default_value_var': 'ExcelのOCR既定値を使用',
            'generate_tif_var': '同名TIFを生成', 'create_tar_var': 'TARを生成',
            'tar_only_var': 'TARだけ残す', 'split_var': 'FORM_IDごとに分割',
            'overwrite_var': '既存出力を上書き',
            'package_include_back_var': '追加時に背面画像を生成',
            'package_result_var': '追加時の背面OCR値',
            'package_scan_batch_id_var': '追加時のCSV位置1',
            'package_arrival_date_var': '追加時のCSV位置4',
            'package_form_id_var': '追加時のCSV位置5',
            'package_application_number_var': '追加時のCSV位置6',
            'package_reception_number_var': '追加時のCSV位置7',
            'package_format_id_var': '追加時のCSV位置8',
            'package_delivery_date_var': '追加時のCSV位置9',
            'package_delivery_shot_var': '追加時のCSV位置10',
        }
        # 変数に対応するラベルを画面から採取し、設定キーは安定した識別子として残す。
        def visit(widget):
            for child in widget.winfo_children():
                try:
                    variable = str(child.cget('textvariable'))
                except Exception:
                    variable = ''
                if variable:
                    try:
                        grid = child.grid_info()
                        column, row = int(grid['column']), int(grid['row'])
                        labels = child.master.grid_slaves(row=row, column=column - 1)
                        label = labels[0].cget('text') if labels else ''
                    except Exception:
                        label = ''
                    for key, value in vars(self).items():
                        if key.endswith('_var') and str(value) == variable and label:
                            descriptions[key] = label
                visit(child)
        visit(self.root)
        return descriptions

    def _load_template(self):
        from .layout_txt_gui import (SETTING_VARIABLE_NAMES, PROFILE_LABELS,
                                     DATE_MODE_LABELS, CSV_STYLE_LABELS)
        if self.template_path is None or not self.template_box.get():
            return
        try:
            settings, fields, packages = load_template(self.template_path, self.template_box.get())
            allowed = set(SETTING_VARIABLE_NAMES) | {'form_var'}
            if set(settings) - allowed or set(SETTING_VARIABLE_NAMES) - set(settings):
                raise ValueError('テンプレート設定キーに不足または未対応の項目があります。')
            for key, value in settings.items():
                variable = getattr(self, key)
                if isinstance(variable.get(), bool) != isinstance(value, bool):
                    raise ValueError('設定の型が一致しません: %s' % key)
            for key, choices in (
                    ('profile_var', [v[0] for v in PROFILE_LABELS]),
                    ('date_mode_var', [v[0] for v in DATE_MODE_LABELS]),
                    ('manifest_style_var', [v[0] for v in CSV_STYLE_LABELS]),
                    ('encoding_var', ['cp932', 'utf-8-sig', 'utf-8']),
                    ('manifest_encoding_var', ['cp932', 'utf-8-sig', 'utf-8'])):
                if settings[key] not in choices:
                    raise ValueError('%s の設定値が不正です: %s' % (key, settings[key]))
        except Exception as exc:
            messagebox.showerror('テンプレート読込エラー', str(exc), parent=self.root)
            return
        if not messagebox.askyesno('テンプレート読込',
                                  '画面の編集内容と出力リストをテンプレートに置き換えますか？', parent=self.root):
            return
        self._finish_cell_edit(save=False)
        self._finish_package_cell_edit(save=False)
        for key, value in settings.items():
            getattr(self, key).set(value)
        self.loaded_fields = [field for field, _include in fields]
        self.form_fields = OrderedDict()
        self.editor_states = {}
        self.displayed_form = None
        for field, _include in fields:
            self.form_fields.setdefault(field.form_id, []).append(field)
        self.form_box.configure(values=list(self.form_fields))
        for form in self.form_fields:
            self.form_var.set(form)
            self._show_selected_form()
            for field, include in fields:
                if field.form_id == form:
                    self.tree.set(field.instance_key, 'include', include)
            self._capture_editor()
        self.form_var.set(settings.get('form_var') if settings.get('form_var') in self.form_fields
                          else next(iter(self.form_fields), ''))
        if self.form_fields:
            self._show_selected_form()
        else:
            self._clear_tree()
        for iid in self.package_tree.get_children(''):
            self.package_tree.delete(iid)
        self.package_grid.reset()
        self.package_items.clear()
        self.package_sequence = 0
        for item, include in packages:
            self.package_sequence += 1
            iid = 'package_%05d' % self.package_sequence
            self.package_items[iid] = item
            self.package_tree.insert('', 'end', iid=iid, values=self._package_values(item, include))
        self.template_mode = True
        self.status_var.set('テンプレート読込完了: %s / %d画像。出力OCRの編集はExcelで変更後、再読込してください。'
                            % (self.template_box.get(), len(packages)))

    def _delete_template(self):
        name = self.template_box.get()
        if self.template_path is None or not name:
            return
        if not messagebox.askyesno('テンプレート削除', '%s を削除しますか？' % name, parent=self.root):
            return
        try:
            delete_template(self.template_path, name)
            self.template_box.configure(values=list_templates(self.template_path))
            self.template_box.set('')
            self.status_var.set('テンプレートを削除しました。画像資材は保管されています。')
        except Exception as exc:
            messagebox.showerror('テンプレート削除エラー', str(exc), parent=self.root)
