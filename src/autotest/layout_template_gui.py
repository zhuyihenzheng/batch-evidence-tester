# -*- coding: utf-8 -*-
"""Excelテンプレートの画面操作。"""

from collections import OrderedDict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .layout_templates import (list_templates, save_template, load_template,
                               scan_template_folder, export_template_batch)


class TemplateGuiMixin(object):
    def _build_template_controls(self, parent):
        from .layout_txt_gui import _wrap_controls
        self.template_path = None
        self.template_name = ''
        self.template_legacy = False
        self.template_mode = False
        self.editor_states = {}
        self.displayed_form = None
        frame = ttk.Frame(parent)
        frame.grid(row=5, column=0, columnspan=8, sticky='we', pady=(6, 0))
        ttk.Button(frame, text='テンプレートを開く...', command=self._choose_template).pack(side='left')
        ttk.Button(frame, text='フォルダから選択...', command=self._choose_template_folder).pack(side='left')
        for text, command in (
                ('再読込', self._load_template), ('名前を付けて保存...', self._new_template),
                ('保存', self._save_template)):
            ttk.Button(frame, text=text, command=command).pack(side='left', padx=2)
        self.template_label = ttk.Label(frame, text='テンプレート: 未選択')
        self.template_label.pack(side='left', padx=5)
        _wrap_controls(frame)

    def _choose_template_folder(self):
        chosen = filedialog.askdirectory(parent=self.root, title='テンプレートフォルダを選択')
        if chosen:
            self._show_template_folder(Path(chosen))

    def _show_template_folder(self, folder):
        import tkinter as tk
        from .layout_txt_gui import _wrap_controls
        try:
            entries, errors = scan_template_folder(folder)
        except Exception as exc:
            messagebox.showerror('フォルダ読込エラー', str(exc), parent=self.root)
            return
        previous = getattr(self, 'template_browser', None)
        if previous is not None and previous.winfo_exists():
            previous.destroy()
        window = tk.Toplevel(self.root)
        self.template_browser = window
        window.title('テンプレート一覧 — Ctrl/Shiftで複数選択')
        window.geometry('%dx%d' % (min(850, self.root.winfo_screenwidth() - 40),
                                   min(520, self.root.winfo_screenheight() - 100)))
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        ttk.Label(window, text=str(folder), wraplength=600).grid(
            row=0, column=0, sticky='we', padx=8, pady=8)
        table_frame = ttk.Frame(window)
        table_frame.grid(row=1, column=0, sticky='nsew', padx=8)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(table_frame, columns=('file', 'name'), show='headings',
                            selectmode='extended')
        self.template_list = tree
        self.template_entries = entries
        tree.heading('file', text='テンプレートExcel')
        tree.heading('name', text='テンプレート名（旧形式は複数行）')
        tree.column('file', width=360, minwidth=120)
        tree.column('name', width=260, minwidth=120)
        tree.grid(row=0, column=0, sticky='nsew')
        bar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        bar.grid(row=0, column=1, sticky='ns')
        tree.configure(yscrollcommand=bar.set)
        for index, (path, name) in enumerate(entries):
            tree.insert('', 'end', iid=str(index), values=(path.name, name))
        controls = ttk.Frame(window)
        controls.grid(row=2, column=0, sticky='we', padx=8, pady=8)
        for label, command in (
                ('全選択', lambda: tree.selection_set(tree.get_children(''))),
                ('選択解除', lambda: tree.selection_remove(tree.selection())),
                ('1件を編集', self._open_selected_template),
                ('選択TXT一括出力', lambda: self._export_selected_templates('txt')),
                ('選択CSV一括出力', lambda: self._export_selected_templates('csv')),
                ('再読込', lambda: self._show_template_folder(folder))):
            ttk.Button(controls, text=label, command=command).pack(side='left', padx=2)
        _wrap_controls(controls)
        tree.bind('<Double-1>', lambda _event: self._open_selected_template())
        ttk.Label(window, text='%d件 / 読込対象外%dファイル。選択した保存済みテンプレートを別々のフォルダへ出力します。'
                  % (len(entries), len(errors)), wraplength=600).grid(
                      row=3, column=0, sticky='we', padx=8, pady=(0, 8))
        if errors:
            messagebox.showwarning('読込対象外のExcel', '\n'.join(
                '%s: %s' % (path.name, reason) for path, reason in errors), parent=window)

    def _selected_templates(self):
        selected = set(self.template_list.selection())
        return [self.template_entries[int(iid)] for iid in self.template_list.get_children('')
                if iid in selected]

    def _open_selected_template(self):
        entries = self._selected_templates()
        if len(entries) != 1:
            messagebox.showwarning('テンプレート選択', '編集するテンプレートを1件選択してください。\n'
                                   '複数選択時は一括出力できます。', parent=self.template_browser)
            return
        self._load_template(path=entries[0][0], name=entries[0][1])

    def _export_selected_templates(self, mode):
        entries = self._selected_templates()
        if not entries:
            messagebox.showwarning('テンプレート選択', 'テンプレートを選択してください。', parent=self.template_browser)
            return
        output = filedialog.askdirectory(parent=self.template_browser, title='一括出力先を選択')
        if not output:
            return
        self.status_var.set('%d件の保存済みテンプレートを一括出力中...' % len(entries))
        self.root.update_idletasks()
        try:
            result = export_template_batch(entries, output, mode)
        except Exception as exc:
            self.status_var.set('一括出力できませんでした: %s' % exc)
            messagebox.showerror('テンプレート一括出力エラー', str(exc), parent=self.template_browser)
            return
        self.status_var.set('一括出力完了: %d件 / %s' % (len(entries), result))
        messagebox.showinfo('テンプレート一括出力完了', '%d件を個別フォルダへ出力しました。\n\n%s'
                            % (len(entries), result), parent=self.template_browser)

    def _choose_template(self):
        chosen = filedialog.askopenfilename(parent=self.root, title='テンプレートExcelを選択',
                                           filetypes=(('Excelテンプレート', '*.xlsx'),))
        if not chosen:
            return
        try:
            names = list_templates(chosen)
            if not names:
                raise ValueError('テンプレートがありません。')
        except Exception as exc:
            messagebox.showerror('テンプレート読込エラー', str(exc), parent=self.root)
            return
        name = names[0]
        if len(names) > 1:
            from tkinter import simpledialog
            number = simpledialog.askinteger(
                '旧形式テンプレートの読込',
                '読み込むテンプレートの番号を入力してください。\n'
                '読込後「名前を付けて保存」で個別ファイルにできます。\n\n' +
                '\n'.join('%d: %s' % (index, value) for index, value in enumerate(names, 1)),
                minvalue=1, maxvalue=len(names), parent=self.root)
            if number is None:
                return
            name = names[number - 1]
        self._load_template(path=Path(chosen), name=name)

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
        self._save_template(save_as=True)

    def _save_template(self, name=None, save_as=False):
        from .layout_txt_gui import SETTING_VARIABLE_NAMES
        target = self.template_path
        if target is None or save_as or self.template_legacy:
            chosen = filedialog.asksaveasfilename(
                parent=self.root, title='1テンプレート1ファイルで保存',
                defaultextension='.xlsx', initialfile=(
                    self.template_path.name if self.template_path and not self.template_legacy
                    else '出力テンプレート.xlsx'),
                filetypes=(('Excelテンプレート', '*.xlsx'),))
            if not chosen:
                return
            target = Path(chosen)
        name = name or target.stem
        overwrite = target.exists()
        if overwrite and not messagebox.askyesno(
                'テンプレート上書き', '%s を現在の設定で上書きしますか？' % target, parent=self.root):
            return
        try:
            self._finish_package_cell_edit(save=True)
            from .layout_txt import coordinate_dimensions
            coordinate_dimensions(self.coordinate_width_var.get(), self.coordinate_height_var.get())
            fields = self._template_fields()
            if not fields and not self.package_items:
                raise ValueError('Excel定義を読み込むか、出力リストへ画像を追加してください。')
            settings = {key: getattr(self, key).get() for key in SETTING_VARIABLE_NAMES}
            settings['form_var'] = self.form_var.get()
            # Excel内の説明列は手編集時の案内として使う。
            descriptions = self._setting_descriptions()
            save_template(target, name, settings, fields,
                          [(self.package_items[iid], self.package_tree.set(iid, 'include'))
                           for iid in self.package_tree.get_children('')],
                          descriptions=descriptions, overwrite=overwrite, single_file=True)
            self.template_path, self.template_name = target, name
            self.template_legacy = False
            self.template_label.configure(text='テンプレート: ' + target.name)
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
            'coordinate_width_var': '自動座標の幅', 'coordinate_height_var': '自動座標の高さ',
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

    def _load_template(self, path=None, name=None):
        from .layout_txt_gui import (SETTING_VARIABLE_NAMES, PROFILE_LABELS,
                                     DATE_MODE_LABELS, CSV_STYLE_LABELS)
        path = Path(path) if path is not None else self.template_path
        if path is None:
            return
        try:
            names = list_templates(path)
            name = name or (names[0] if len(names) == 1 else self.template_name)
            if name not in names:
                raise ValueError('テンプレートを開き直してください。')
            settings, fields, packages = load_template(path, name)
            # 旧テンプレートは従来の50×50として読み込み、現在の画面値を引き継がない。
            settings.setdefault('coordinate_width_var', '50')
            settings.setdefault('coordinate_height_var', '50')
            from .layout_txt import coordinate_dimensions
            coordinate_dimensions(settings['coordinate_width_var'], settings['coordinate_height_var'])
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
        self.template_path, self.template_name = path, name
        self.template_legacy = len(names) > 1
        self.template_label.configure(text='テンプレート: ' + path.name)
        self.status_var.set('テンプレート読込完了: %s / %d画像。出力リストを選択して「選択行の内容を編集」でOCRを変更できます。'
                            % (path.name, len(packages)))
