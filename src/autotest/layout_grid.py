# -*- coding: utf-8 -*-
"""Treeviewに矩形セル選択とExcel互換のTSV編集を追加する。"""

import csv
import io


def parse_clipboard(text):
    if text == '':
        return [['']]
    try:
        rows = list(csv.reader(io.StringIO(text), delimiter='\t', strict=True))
    except csv.Error as exc:
        raise ValueError('貼り付けデータを読み取れません: %s' % exc)
    rows = [row or [''] for row in rows]
    if not rows or any(len(row) != len(rows[0]) for row in rows):
        raise ValueError('貼り付けデータの列数が揃っていません。')
    return rows


def clipboard_text(rows):
    stream = io.StringIO()
    csv.writer(stream, delimiter='\t', lineterminator='\r\n').writerows(rows)
    return stream.getvalue()


def paste_changes(rows, columns, bounds, values, editable):
    r0, c0, r1, c1 = bounds
    height, width = len(values), len(values[0])
    selected_height, selected_width = r1 - r0 + 1, c1 - c0 + 1
    if selected_height > 1 or selected_width > 1:
        if selected_height % height or selected_width % width:
            raise ValueError('選択範囲と貼り付けデータの大きさが一致しません。')
        height, width = selected_height, selected_width
    if r0 + height > len(rows) or c0 + width > len(columns):
        raise ValueError('貼り付け先が表の範囲外です。行・列は自動追加しません。')
    changes = {}
    for r in range(height):
        for c in range(width):
            column = columns[c0 + c]
            if column not in editable:
                raise ValueError('読取専用列を含むため貼り付けできません: %s' % column)
            changes[(rows[r0 + r], column)] = values[r % len(values)][c % len(values[0])]
    return changes, (r0, c0, r0 + height - 1, c0 + width - 1)


class CellGrid(object):
    def __init__(self, tree, editable, apply, status):
        import tkinter as tk
        self.tree, self.editable, self.apply, self.status = tree, set(editable), apply, status
        self.anchor = self.active = None
        self.dragging = False
        self.editor = None
        self.undo_stack, self.redo_stack = [], []
        self.borders = [tk.Frame(tree, background='#1674c4') for _ in range(4)]
        self.menu = tk.Menu(tree, tearoff=False)
        for text, command in (('コピー', self.copy), ('切り取り', self.cut),
                              ('貼り付け', self.paste), ('内容をクリア', self.clear),
                              ('元に戻す', self.undo), ('やり直し', self.redo)):
            self.menu.add_command(label=text, command=command)
        tree.bind('<Button-1>', self.click)
        tree.bind('<B1-Motion>', self.drag)
        tree.bind('<ButtonRelease-1>', self.release)
        tree.bind('<Double-1>', self.edit)
        tree.bind('<Button-3>', self.popup)
        tree.bind('<KeyPress>', self.key)
        tree.bind('<Configure>', lambda _e: self.draw(), add='+')
        for option in ('xscrollcommand', 'yscrollcommand'):
            original = tree.cget(option)
            def scroll(first, last, command=original):
                self.finish(True)
                if command:
                    tree.tk.call(*tree.tk.splitlist(command), first, last)
                self.draw()
            tree.configure(**{option: scroll})

    def columns(self):
        columns = self.tree.cget('displaycolumns')
        if isinstance(columns, str):
            columns = self.tree.tk.splitlist(columns)
        return tuple(self.tree.cget('columns')) if tuple(columns) == ('#all',) else tuple(columns)

    def reset(self):
        self.finish(False)
        self.anchor = self.active = None
        self.undo_stack, self.redo_stack = [], []
        self.draw()

    def bounds(self):
        rows, columns = self.tree.get_children(''), self.columns()
        if (not self.anchor or not self.active or self.anchor[0] not in rows or
                self.active[0] not in rows or self.anchor[1] not in columns or self.active[1] not in columns):
            return None
        a, b = rows.index(self.anchor[0]), rows.index(self.active[0])
        c, d = columns.index(self.anchor[1]), columns.index(self.active[1])
        return min(a, b), min(c, d), max(a, b), max(c, d)

    def hit(self, event):
        row, token = self.tree.identify_row(event.y), self.tree.identify_column(event.x)
        if self.tree.identify_region(event.x, event.y) != 'cell' or not row or not token:
            return None
        index = int(token[1:]) - 1
        columns = self.columns()
        return (row, columns[index]) if 0 <= index < len(columns) else None

    def select(self, cell, extend=False):
        if not extend or self.anchor is None:
            self.anchor = cell
        self.active = cell
        bounds = self.bounds()
        if bounds:
            rows = self.tree.get_children('')
            self.tree.selection_set(rows[bounds[0]:bounds[2] + 1])
            self.tree.focus(cell[0])
            self.status('%d行 × %d列を選択 / Ctrl+C コピー・Ctrl+V 貼り付け・Delete クリア'
                        % (bounds[2] - bounds[0] + 1, bounds[3] - bounds[1] + 1))
        self.draw()

    def click(self, event):
        self.dragging = False
        cell = self.hit(event)
        if cell:
            if not self.finish(True):
                return 'break'
            self.tree.focus_set()
            self.dragging = True
            self.select(cell, bool(event.state & 1))
            return 'break'

    def drag(self, event):
        if self.anchor and self.dragging:
            if event.y < 25:
                self.tree.yview_scroll(-1, 'units')
            elif event.y > self.tree.winfo_height() - 8:
                self.tree.yview_scroll(1, 'units')
            if event.x < 8:
                self.tree.xview_scroll(-1, 'units')
            elif event.x > self.tree.winfo_width() - 8:
                self.tree.xview_scroll(1, 'units')
            cell = self.hit(event)
            if cell:
                self.select(cell, True)
            return 'break'

    def release(self, _event):
        self.dragging = False

    def draw(self):
        for border in self.borders:
            border.place_forget()
        bounds = self.bounds()
        if not bounds:
            return
        rows, columns = self.tree.get_children(''), self.columns()
        boxes = [self.tree.bbox(row, columns[bounds[1]])
                 for row in rows[bounds[0]:bounds[2] + 1]]
        boxes = [box for box in boxes if box]
        if not boxes:
            return
        end = self.tree.bbox(rows[bounds[0]], columns[bounds[3]])
        if not end:
            end = self.tree.bbox(next(row for row in rows[bounds[0]:bounds[2] + 1]
                                      if self.tree.bbox(row)), columns[bounds[3]])
        if not end:
            return
        left, top = max(0, boxes[0][0]), boxes[0][1]
        right = min(self.tree.winfo_width(), end[0] + end[2])
        bottom = min(self.tree.winfo_height(), boxes[-1][1] + boxes[-1][3])
        if right <= left or bottom <= top:
            return
        for border, geometry in zip(self.borders, (
                (left, top, right-left, 2), (left, bottom-2, right-left, 2),
                (left, top, 2, bottom-top), (right-2, top, 2, bottom-top))):
            border.place(x=geometry[0], y=geometry[1], width=geometry[2], height=geometry[3])

    def report(self, exc):
        from tkinter import messagebox
        self.status(str(exc))
        messagebox.showerror('セル編集エラー', str(exc), parent=self.tree.winfo_toplevel())

    def change(self, changes):
        if not changes:
            return True
        try:
            reverse = self.apply(changes)
        except (ValueError, RuntimeError) as exc:
            self.report(exc)
            return False
        self.undo_stack.append(reverse)
        self.undo_stack = self.undo_stack[-50:]
        self.redo_stack = []
        self.draw()
        return True

    def copy(self):
        if not self.finish(True):
            return False
        bounds = self.bounds()
        if not bounds:
            return False
        r0, c0, r1, c1 = bounds
        rows, columns = self.tree.get_children(''), self.columns()
        text = clipboard_text([[self.tree.set(row, col) for col in columns[c0:c1+1]]
                               for row in rows[r0:r1+1]])
        self.tree.clipboard_clear()
        self.tree.clipboard_append(text)
        return True

    def cut(self):
        if self.copy():
            self.clear()

    def clear(self):
        if not self.finish(True):
            return
        bounds = self.bounds()
        if bounds:
            r0, c0, r1, c1 = bounds
            self.change({(row, col): '0' if col == 'include' else ''
                         for row in self.tree.get_children('')[r0:r1+1]
                         for col in self.columns()[c0:c1+1] if col in self.editable})

    def paste(self):
        if not self.finish(True) or not self.bounds():
            return
        try:
            values = parse_clipboard(self.tree.clipboard_get())
            changes, bounds = paste_changes(self.tree.get_children(''), self.columns(),
                                            self.bounds(), values, self.editable)
        except Exception as exc:
            self.report(exc)
            return
        if self.change(changes):
            rows, columns = self.tree.get_children(''), self.columns()
            self.anchor = (rows[bounds[0]], columns[bounds[1]])
            self.select((rows[bounds[2]], columns[bounds[3]]), True)

    def undo(self):
        if self.finish(True) and self.undo_stack:
            self.redo_stack.append(self.undo_stack.pop()())
            self.draw()

    def redo(self):
        if self.finish(True) and self.redo_stack:
            self.undo_stack.append(self.redo_stack.pop()())
            self.draw()

    def popup(self, event):
        cell = self.hit(event)
        if cell:
            bounds = self.bounds()
            rows, columns = self.tree.get_children(''), self.columns()
            if not bounds or not (bounds[0] <= rows.index(cell[0]) <= bounds[2] and
                                  bounds[1] <= columns.index(cell[1]) <= bounds[3]):
                self.select(cell)
            self.tree.focus_set()
            try:
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.menu.grab_release()
        return 'break'

    def move(self, dr, dc, extend=False):
        rows, columns = self.tree.get_children(''), self.columns()
        if not rows or not columns:
            return
        if not self.bounds():
            self.select((rows[0], columns[0]))
            return
        r, c = rows.index(self.active[0]), columns.index(self.active[1])
        r, c = max(0, min(len(rows)-1, r+dr)), max(0, min(len(columns)-1, c+dc))
        self.tree.see(rows[r])
        self.select((rows[r], columns[c]), extend)
        box = self.tree.bbox(rows[r], columns[c])
        if box:
            total = sum(int(self.tree.column(col, 'width')) for col in columns)
            offset = self.tree.xview()[0] * total
            if box[0] < 0:
                self.tree.xview_moveto(max(0, offset + box[0]) / max(1, total))
            elif box[0] + box[2] > self.tree.winfo_width():
                self.tree.xview_moveto((offset + box[0] + box[2] - self.tree.winfo_width()) / max(1, total))

    def key(self, event):
        key = event.keysym.lower()
        command = bool(event.state & 4) or (self.tree.tk.call('tk', 'windowingsystem') == 'aqua' and bool(event.state & 8))
        if command:
            action = {'c': self.copy, 'x': self.cut, 'v': self.paste,
                      'z': self.undo, 'y': self.redo}.get(key)
            if key == 'a':
                rows, columns = self.tree.get_children(''), self.columns()
                if rows and columns:
                    self.anchor = (rows[0], columns[0])
                    self.select((rows[-1], columns[-1]), True)
            elif action:
                action()
            return 'break'
        if key in ('delete', 'backspace'):
            self.clear()
        elif key in ('up', 'down', 'left', 'right', 'return', 'tab', 'iso_left_tab'):
            dr, dc = {'up': (-1, 0), 'down': (1, 0), 'left': (0, -1), 'right': (0, 1),
                      'return': (-1 if event.state & 1 else 1, 0),
                      'tab': (0, -1 if event.state & 1 else 1), 'iso_left_tab': (0, -1)}[key]
            self.move(dr, dc, bool(event.state & 1) and key in ('up', 'down', 'left', 'right'))
        elif key == 'f2':
            self.edit()
        elif event.char and event.char.isprintable():
            self.edit(initial=event.char)
        else:
            return
        return 'break'

    def edit(self, event=None, initial=None):
        from tkinter import ttk
        if event is not None:
            cell = self.hit(event)
            if cell:
                self.select(cell)
        if not self.bounds() or self.active[1] not in self.editable:
            return 'break'
        if not self.finish(True):
            return 'break'
        cell = self.active
        if cell[1] == 'include' and initial is None:
            self.change({cell: '0' if self.tree.set(*cell) == '1' else '1'})
            return 'break'
        box = self.tree.bbox(*cell)
        if not box:
            return 'break'
        self.edit_cell = cell
        self.editor = ttk.Entry(self.tree)
        self.editor.insert(0, self.tree.set(*cell) if initial is None else initial)
        if initial is None:
            self.editor.select_range(0, 'end')
        self.editor.place(x=box[0], y=box[1], width=box[2], height=box[3])
        self.editor.focus_set()
        for sequence, dr, dc in (('<Return>', 1, 0), ('<Shift-Return>', -1, 0),
                                  ('<Tab>', 0, 1), ('<Shift-Tab>', 0, -1)):
            self.editor.bind(sequence, lambda _e, r=dr, c=dc: self.end_edit(r, c))
        if self.tree.tk.call('tk', 'windowingsystem') == 'x11':
            self.editor.bind('<ISO_Left_Tab>', lambda _e: self.end_edit(0, -1))
        self.editor.bind('<Escape>', lambda _e: self.end_edit(0, 0, False))
        self.editor.bind('<FocusOut>', lambda _e: self.finish(True))
        self.editor.bind('<<Paste>>', self.editor_paste)
        self.editor.bind('<Control-v>', self.editor_paste)
        if self.tree.tk.call('tk', 'windowingsystem') == 'aqua':
            self.editor.bind('<Command-v>', self.editor_paste)
        return 'break'

    def editor_paste(self, _event=None):
        try:
            values = parse_clipboard(self.tree.clipboard_get())
        except Exception:
            return
        if len(values) > 1 or len(values[0]) > 1:
            self.paste()
            self.tree.focus_set()
            return 'break'

    def finish(self, save=True):
        if self.editor is None:
            return True
        editor, cell = self.editor, self.edit_cell
        self.editor = None
        value = editor.get()
        editor.destroy()
        return not save or self.change({cell: value})

    def end_edit(self, dr, dc, save=True):
        if self.finish(save):
            self.tree.focus_set()
            self.move(dr, dc)
        return 'break'
