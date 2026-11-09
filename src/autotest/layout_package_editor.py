# -*- coding: utf-8 -*-
"""出力リスト内の認識TXTを、原本を変更せず下書きとして編集する。"""

import csv
import io

from .layout_grid import CellGrid
from .layout_tar import LayoutTarError, _form_records


COLUMNS = ('field_id', 'value', 'attribute_flag', 'coordinates')


class PackageContentEditor(object):
    def __init__(self, parent, text, title, on_save):
        import tkinter as tk
        from tkinter import ttk
        self.records = _form_records(text)
        if not self.records:
            raise LayoutTarError('編集する正面TXTがありません。FORMを追加してください。')
        self.original = [list(row) for row in self.records]
        self.on_save = on_save
        self.index = 0
        self.window = tk.Toplevel(parent)
        self.window.title('出力リストの内容を編集 — ' + title)
        self.window.geometry('%dx%d' % (min(900, parent.winfo_screenwidth() - 40),
                                       min(600, parent.winfo_screenheight() - 100)))
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(2, weight=1)
        self.form_var = tk.StringVar(master=self.window)
        self.presence_var = tk.StringVar(master=self.window)
        self.status_var = tk.StringVar(master=self.window)
        top = ttk.Frame(self.window, padding=8)
        top.grid(row=0, column=0, sticky='we')
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text='TXT内の帳票:').grid(row=0, column=0, sticky='w')
        self.form_box = ttk.Combobox(top, state='readonly')
        self.form_box.grid(row=0, column=1, sticky='we', padx=6)
        self.form_box.bind('<<ComboboxSelected>>', self.switch)
        fields = ttk.Frame(self.window, padding=(8, 0))
        fields.grid(row=1, column=0, sticky='we')
        ttk.Label(fields, text='FORM_ID:').grid(row=0, column=0, sticky='w')
        ttk.Entry(fields, textvariable=self.form_var, width=16).grid(row=0, column=1, padx=6)
        ttk.Label(fields, text='対象有無:').grid(row=0, column=2, sticky='w')
        ttk.Entry(fields, textvariable=self.presence_var, width=10).grid(row=0, column=3, padx=6)
        table = ttk.Frame(self.window, padding=8)
        table.grid(row=2, column=0, sticky='nsew')
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=COLUMNS, show='headings', selectmode='extended')
        for column, heading, width in zip(COLUMNS, ('FIELD_ID', 'OCR値', '属性', '座標（X,Y,幅,高さ）'),
                                           (100, 350, 80, 200)):
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, minwidth=50, stretch=column == 'value')
        self.tree.grid(row=0, column=0, sticky='nsew')
        ybar = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        xbar = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        ybar.grid(row=0, column=1, sticky='ns')
        xbar.grid(row=1, column=0, sticky='we')
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.grid = CellGrid(self.tree, COLUMNS, self.apply, self.status_var.set)
        ttk.Label(self.window, textvariable=self.status_var, wraplength=600).grid(
            row=3, column=0, sticky='we', padx=8)
        ttk.Label(self.window,
                  text='範囲選択・コピー・貼付・Delete・Ctrl+Zに対応。保存で出力リストのTXTを更新します。\n'
                       '画像・CSV位置5は変更しません。再利用する場合はメイン画面でテンプレートを保存してください。',
                  wraplength=600).grid(row=4, column=0, sticky='we', padx=8, pady=6)
        buttons = ttk.Frame(self.window, padding=8)
        buttons.grid(row=5, column=0, sticky='e')
        ttk.Button(buttons, text='キャンセル', command=self.cancel).pack(side='right', padx=4)
        ttk.Button(buttons, text='出力リストへ保存', command=self.save).pack(side='right', padx=4)
        self.window.protocol('WM_DELETE_WINDOW', self.cancel)
        self.show()

    def show(self):
        self.grid.reset()
        for iid in self.tree.get_children(''):
            self.tree.delete(iid)
        record = self.records[self.index]
        self.form_var.set(record[0])
        self.presence_var.set(record[1])
        self.form_box.configure(values=['%d: FORM_ID %s' % (i + 1, row[0])
                                        for i, row in enumerate(self.records)])
        self.form_box.current(self.index)
        for offset in range(2, len(record), 4):
            self.tree.insert('', 'end', iid=str((offset - 2) // 4), values=record[offset:offset + 4])
        self.status_var.set('帳票 %d/%d / %d項目' % (self.index + 1, len(self.records), (len(record) - 2) // 4))

    def restore(self, snapshot):
        previous = {iid: tuple(self.tree.item(iid, 'values')) for iid in snapshot}
        for iid, values in snapshot.items():
            self.tree.item(iid, values=values)
        return lambda: self.restore(previous)

    def apply(self, changes):
        pending = {}
        for (iid, column), value in changes.items():
            if not self.tree.exists(iid) or column not in COLUMNS:
                raise ValueError('編集できないセルです。')
            row = pending.setdefault(iid, list(self.tree.item(iid, 'values')))
            row[COLUMNS.index(column)] = value
        for row in pending.values():
            if row[1] == '':
                row[2] = row[3] = ''
        return self.restore(pending)

    def draft(self):
        row = [self.form_var.get(), self.presence_var.get()]
        for iid in self.tree.get_children(''):
            values = list(self.tree.item(iid, 'values'))
            if values[1] == '':
                values[2] = values[3] = ''
            row.extend(values)
        return row

    def commit(self):
        from tkinter import messagebox
        if not self.grid.finish(True):
            return False
        row = self.draft()
        if not row[0].isdigit():
            messagebox.showerror('編集エラー', 'FORM_IDは数字を入力してください。', parent=self.window)
            return False
        self.records[self.index] = row
        return True

    def switch(self, _event=None):
        selected = self.form_box.current()
        if not self.commit():
            self.form_box.current(self.index)
            return
        self.index = selected
        self.show()

    def save(self):
        from tkinter import messagebox
        if not self.commit():
            return
        stream = io.StringIO()
        writer = csv.writer(stream, quoting=csv.QUOTE_ALL, lineterminator='\r\n')
        writer.writerows(self.records)
        try:
            self.on_save(stream.getvalue())
        except (LayoutTarError, ValueError) as exc:
            messagebox.showerror('保存エラー', str(exc), parent=self.window)
            return
        self.window.destroy()

    def cancel(self):
        from tkinter import messagebox
        if (self.grid.editor is not None or self.records != self.original or
                self.draft() != self.records[self.index]):
            if not messagebox.askyesno('編集を破棄', '保存せずに閉じますか？', parent=self.window):
                return
        self.grid.reset()
        self.window.destroy()
