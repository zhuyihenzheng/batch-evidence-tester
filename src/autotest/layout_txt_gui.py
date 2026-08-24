# -*- coding: utf-8 -*-
"""Excel -> Layout TXT/TIF/TAR 生成ツールの単独GUI。"""

import sys
from collections import OrderedDict
from pathlib import Path
from typing import Optional

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError as exc:  # pragma: no cover - GUI無し環境
    raise SystemExit("画面を開けません（tkinterが見つかりません）: %s" % exc)

from openpyxl import load_workbook

from .layout_txt import (
    LayoutTxtError,
    _print_result,
    generate_layout_txt,
    read_layout_fields,
)


PROFILE_LABELS = (
    ("通常値（推奨）", "normal"),
    ("最大桁数ちょうど", "max"),
    ("最大桁数 + 1（異常系）", "over"),
)

FORMAT_LABELS = (
    ("1帳票1行（全値ダブルクォート）", "raw"),
    ("ラベル付き（確認用）", "labeled"),
    ("TSV（1項目1行）", "tsv"),
)

DATE_MODE_LABELS = (
    ("通常Form=和暦 / 4001=4形式網羅（推奨）", "coverage"),
    ("全Formで4形式を順番使用", "cycle"),
    ("和暦: 5/8/6/1", "wareki"),
    ("西暦: /2026/6/1", "seireki"),
    ("元号+西暦: 5/2026/6/1", "era-seireki"),
    ("複数: 5/8/6/1|5/8/6/2", "multiple"),
)

ERROR_PATTERN_LABELS = (
    ("全24 Pattern（推奨）", "all"),
    ("主要8 Pattern", "core"),
    ("異常Patternなし", "none"),
)

TREE_COLUMNS = (
    "include", "row", "layout_id", "field_id", "item_name",
    "data_type", "ime_name", "max_digits", "value",
    "attribute_flag", "coordinates",
)

EDITABLE_COLUMNS = {
    "include", "field_id", "value", "attribute_flag", "coordinates",
}


def _value_for_label(pairs, label):
    for item_label, value in pairs:
        if item_label == label:
            return value
    return pairs[0][1]


class LayoutTxtGui(object):
    def __init__(self, root, initial_excel: Optional[Path] = None) -> None:
        self.root = root
        root.title("Layout TXT / TIF / TAR 生成ツール")
        root.geometry("1280x900")
        root.minsize(1020, 720)

        self.excel_var = tk.StringVar(value=str(initial_excel or ""))
        initial_output = (initial_excel.parent / "layout_txt") if initial_excel else ""
        self.output_var = tk.StringVar(value=str(initial_output))
        self.sheet_var = tk.StringVar(value="")
        self.header_var = tk.StringVar(value="自動")
        self.form_col_var = tk.StringVar(value="auto")
        self.layout_col_var = tk.StringVar(value="auto")
        self.field_col_var = tk.StringVar(value="auto")
        self.item_col_var = tk.StringVar(value="auto")
        self.type_col_var = tk.StringVar(value="I")
        self.ime_col_var = tk.StringVar(value="J")
        self.max_col_var = tk.StringVar(value="K")
        self.profile_var = tk.StringVar(value=PROFILE_LABELS[0][0])
        self.date_mode_var = tk.StringVar(value=DATE_MODE_LABELS[0][0])
        self.coverage_form_var = tk.StringVar(value="4001")
        self.error_pattern_var = tk.StringVar(value=ERROR_PATTERN_LABELS[0][0])
        self.filename_template_var = tk.StringVar(value="{form_id}")
        self.generate_tif_var = tk.BooleanVar(value=True)
        self.create_tar_var = tk.BooleanVar(value=True)
        self.tar_only_var = tk.BooleanVar(value=False)
        self.tar_name_var = tk.StringVar(value="{source}_layout_data")
        self.format_var = tk.StringVar(value=FORMAT_LABELS[0][0])
        self.encoding_var = tk.StringVar(value="cp932")
        self.split_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.form_var = tk.StringVar(value="")
        self.form_summary_var = tk.StringVar(value="Excel定義を読み込んでください。")
        self.status_var = tk.StringVar(value="Excelを選択してください。")

        self.loaded_fields = []  # type: List[LayoutField]
        self.form_fields = OrderedDict()  # type: OrderedDict[str, List[LayoutField]]
        self.cell_editor = None
        self.editor_context = None

        self._build()
        if initial_excel:
            self._load_sheets()
            self._read_definitions(show_errors=False)

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(5, weight=1)

        ttk.Label(outer, text="Excel → Layout TXT / TIF / TAR",
                  font=("", 16, "bold")).grid(row=0, column=0, sticky="w")

        source = ttk.LabelFrame(outer, text="入力・列設定", padding=8)
        source.grid(row=1, column=0, sticky="we", pady=(8, 0))
        for column in range(8):
            source.columnconfigure(column, weight=1 if column in (1, 3, 5) else 0)

        ttk.Label(source, text="入力Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(source, textvariable=self.excel_var).grid(
            row=0, column=1, columnspan=6, sticky="we", padx=(6, 6))
        ttk.Button(source, text="参照...", command=self._choose_excel).grid(row=0, column=7)

        ttk.Label(source, text="出力先:").grid(row=1, column=0, sticky="w", pady=(7, 0))
        ttk.Entry(source, textvariable=self.output_var).grid(
            row=1, column=1, columnspan=6, sticky="we", padx=(6, 6), pady=(7, 0))
        ttk.Button(source, text="参照...", command=self._choose_output).grid(
            row=1, column=7, pady=(7, 0))

        ttk.Label(source, text="シート:").grid(row=2, column=0, sticky="w", pady=(7, 0))
        self.sheet_box = ttk.Combobox(
            source, textvariable=self.sheet_var, state="readonly", width=24)
        self.sheet_box.grid(row=2, column=1, sticky="we", padx=(6, 14), pady=(7, 0))
        self.sheet_box.bind("<<ComboboxSelected>>", self._definition_setting_changed)
        ttk.Label(source, text="見出し行:").grid(row=2, column=2, sticky="e", pady=(7, 0))
        ttk.Entry(source, textvariable=self.header_var, width=9).grid(
            row=2, column=3, sticky="w", padx=(6, 14), pady=(7, 0))
        ttk.Button(source, text="Excel定義を読込", command=self._read_definitions).grid(
            row=2, column=6, columnspan=2, sticky="e", pady=(7, 0))

        column_items = (
            ("FormID", self.form_col_var), ("LayoutID", self.layout_col_var),
            ("FieldID", self.field_col_var), ("ITEM_NAME", self.item_col_var),
            ("DATA_TYPE", self.type_col_var), ("IME", self.ime_col_var),
            ("MAX", self.max_col_var),
        )
        columns_frame = ttk.Frame(source)
        columns_frame.grid(row=3, column=0, columnspan=8, sticky="we", pady=(7, 0))
        for index, (label, variable) in enumerate(column_items):
            ttk.Label(columns_frame, text=label + ":").grid(row=0, column=index * 2, sticky="w")
            ttk.Entry(columns_frame, textvariable=variable, width=9).grid(
                row=0, column=index * 2 + 1, sticky="w", padx=(3, 10))

        generation = ttk.LabelFrame(outer, text="生成・ファイル設定", padding=8)
        generation.grid(row=2, column=0, sticky="we", pady=(8, 0))
        for column in range(8):
            generation.columnconfigure(column, weight=1 if column in (1, 3, 5, 7) else 0)

        ttk.Label(generation, text="データ:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            generation, textvariable=self.profile_var,
            values=[label for label, _value in PROFILE_LABELS],
            state="readonly", width=19).grid(row=0, column=1, sticky="we", padx=(5, 12))
        ttk.Label(generation, text="日付:").grid(row=0, column=2, sticky="w")
        ttk.Combobox(
            generation, textvariable=self.date_mode_var,
            values=[label for label, _value in DATE_MODE_LABELS],
            state="readonly", width=27).grid(row=0, column=3, sticky="we", padx=(5, 12))
        ttk.Label(generation, text="全網羅Form:").grid(row=0, column=4, sticky="w")
        ttk.Entry(generation, textvariable=self.coverage_form_var, width=10).grid(
            row=0, column=5, sticky="w", padx=(5, 12))
        ttk.Label(generation, text="全出力Pattern:").grid(row=0, column=6, sticky="w")
        ttk.Combobox(
            generation, textvariable=self.error_pattern_var,
            values=[label for label, _value in ERROR_PATTERN_LABELS],
            state="readonly", width=18).grid(row=0, column=7, sticky="we", padx=(5, 0))

        ttk.Label(generation, text="TXT名:").grid(row=1, column=0, sticky="w", pady=(7, 0))
        ttk.Entry(generation, textvariable=self.filename_template_var).grid(
            row=1, column=1, sticky="we", padx=(5, 12), pady=(7, 0))
        ttk.Label(generation, text="TAR名:").grid(row=1, column=2, sticky="w", pady=(7, 0))
        ttk.Entry(generation, textvariable=self.tar_name_var).grid(
            row=1, column=3, sticky="we", padx=(5, 12), pady=(7, 0))
        ttk.Label(generation, text="形式:").grid(row=1, column=4, sticky="w", pady=(7, 0))
        ttk.Combobox(
            generation, textvariable=self.format_var,
            values=[label for label, _value in FORMAT_LABELS],
            state="readonly", width=16).grid(
                row=1, column=5, sticky="we", padx=(5, 12), pady=(7, 0))
        ttk.Label(generation, text="文字コード:").grid(row=1, column=6, sticky="w", pady=(7, 0))
        ttk.Combobox(
            generation, textvariable=self.encoding_var,
            values=("cp932", "utf-8-sig", "utf-8"),
            state="readonly", width=12).grid(
                row=1, column=7, sticky="we", padx=(5, 0), pady=(7, 0))

        flags = ttk.Frame(generation)
        flags.grid(row=2, column=0, columnspan=8, sticky="w", pady=(7, 0))
        ttk.Checkbutton(flags, text="FormIDごとに分割", variable=self.split_var).pack(
            side="left", padx=(0, 14))
        ttk.Checkbutton(flags, text="同名TIFを生成", variable=self.generate_tif_var).pack(
            side="left", padx=(0, 14))
        ttk.Checkbutton(flags, text="TARを生成", variable=self.create_tar_var).pack(
            side="left", padx=(0, 14))
        ttk.Checkbutton(flags, text="TARだけ残す", variable=self.tar_only_var,
                        command=self._tar_only_changed).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(flags, text="既存ファイルを上書き",
                        variable=self.overwrite_var).pack(side="left")

        ttk.Label(
            generation,
            text="TXT名: {form_id}/{pattern}/{seq:02d}/{source}  TAR名: {form_id}/{source}",
            foreground="#666").grid(row=3, column=0, columnspan=8, sticky="w", pady=(5, 0))

        selector = ttk.LabelFrame(outer, text="FORM_IDを選択して内容を編集", padding=8)
        selector.grid(row=3, column=0, sticky="we", pady=(8, 0))
        selector.columnconfigure(5, weight=1)
        ttk.Label(selector, text="FORM_ID:").grid(row=0, column=0, sticky="w")
        self.form_box = ttk.Combobox(
            selector, textvariable=self.form_var, state="readonly", width=18)
        self.form_box.grid(row=0, column=1, sticky="w", padx=(6, 10))
        self.form_box.bind("<<ComboboxSelected>>", self._show_selected_form)
        ttk.Button(selector, text="表示", command=self._show_selected_form).grid(row=0, column=2)
        ttk.Button(selector, text="全項目ON", command=lambda: self._set_all_includes("1")).grid(
            row=0, column=3, padx=(10, 4))
        ttk.Button(selector, text="全項目OFF", command=lambda: self._set_all_includes("0")).grid(
            row=0, column=4, padx=(0, 10))
        ttk.Label(selector, textvariable=self.form_summary_var, foreground="#444").grid(
            row=0, column=5, sticky="w")
        ttk.Label(
            selector,
            text="出力/OCR値/属性/ELEMENT_ID/座標はダブルクリックで編集できます。",
            foreground="#666").grid(row=1, column=0, columnspan=6, sticky="w", pady=(5, 0))

        table_frame = ttk.Frame(outer)
        table_frame.grid(row=5, column=0, sticky="nsew", pady=(8, 0))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            table_frame, columns=TREE_COLUMNS, show="headings", selectmode="browse")
        headings = {
            "include": "出力", "row": "Excel行", "layout_id": "LAYOUT_ID",
            "field_id": "ELEMENT_ID", "item_name": "ITEM_NAME",
            "data_type": "DATA_TYPE", "ime_name": "IME", "max_digits": "MAX",
            "value": "OCR生成値（編集可）", "attribute_flag": "属性",
            "coordinates": "座標",
        }
        widths = {
            "include": 48, "row": 62, "layout_id": 90, "field_id": 100,
            "item_name": 170, "data_type": 120, "ime_name": 110,
            "max_digits": 65, "value": 230, "attribute_flag": 60,
            "coordinates": 120,
        }
        for name in TREE_COLUMNS:
            self.tree.heading(name, text=headings[name])
            self.tree.column(name, width=widths[name], minwidth=45,
                             stretch=name in ("item_name", "value"))
        ybar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="we")
        self.tree.bind("<Double-1>", self._begin_cell_edit)

        status = tk.Label(outer, textvariable=self.status_var, anchor="w", justify="left",
                          fg="#333", wraplength=1160)
        status.grid(row=6, column=0, sticky="we", pady=(8, 4))

        buttons = ttk.Frame(outer)
        buttons.grid(row=7, column=0, sticky="e")
        ttk.Button(buttons, text="閉じる", command=self.root.destroy).pack(
            side="right", padx=(8, 0))
        self.all_button = ttk.Button(
            buttons, text="全FORM_IDを直接出力", command=self._generate_all)
        self.all_button.pack(side="right", padx=(8, 0))
        self.selected_button = ttk.Button(
            buttons, text="表示中FORM_IDを出力", command=self._generate_selected)
        self.selected_button.pack(side="right")

    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root, title="レイアウト定義Excelを選択",
            filetypes=(("Excel", "*.xlsx *.xlsm"), ("すべて", "*.*")))
        if not path:
            return
        self.excel_var.set(path)
        if not self.output_var.get():
            self.output_var.set(str(Path(path).parent / "layout_txt"))
        self._load_sheets()
        self._read_definitions(show_errors=False)

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(parent=self.root, title="出力フォルダを選択")
        if path:
            self.output_var.set(path)

    def _load_sheets(self) -> None:
        path = Path(self.excel_var.get().strip())
        if not path.is_file():
            return
        try:
            wb = load_workbook(str(path), read_only=True, data_only=True)
            try:
                names = list(wb.sheetnames)
                active = wb.active.title
            finally:
                wb.close()
        except Exception as exc:
            messagebox.showerror("Excel読込エラー", str(exc), parent=self.root)
            return
        self.sheet_box.config(values=names)
        self.sheet_var.set(active)

    def _definition_setting_changed(self, _event=None) -> None:
        self.loaded_fields = []
        self.form_fields = OrderedDict()
        self.form_box.config(values=[])
        self.form_var.set("")
        self._clear_tree()
        self.status_var.set("設定が変わりました。「Excel定義を読込」を押してください。")

    def _header_row(self):
        raw = self.header_var.get().strip()
        if not raw or raw in ("自動", "auto", "AUTO"):
            return None
        try:
            value = int(raw)
        except ValueError:
            raise LayoutTxtError("見出し行は数値または「自動」にしてください。")
        if value < 1:
            raise LayoutTxtError("見出し行は1以上にしてください。")
        return value

    def _read_args(self) -> dict:
        return {
            "excel_path": Path(self.excel_var.get().strip()),
            "sheet_name": self.sheet_var.get() or None,
            "header_row": self._header_row(),
            "form_column": self.form_col_var.get(),
            "layout_column": self.layout_col_var.get(),
            "field_column": self.field_col_var.get(),
            "item_column": self.item_col_var.get(),
            "data_type_column": self.type_col_var.get(),
            "ime_column": self.ime_col_var.get(),
            "max_digits_column": self.max_col_var.get(),
            "profile": _value_for_label(PROFILE_LABELS, self.profile_var.get()),
            "date_mode": _value_for_label(DATE_MODE_LABELS, self.date_mode_var.get()),
            "coverage_form_id": self.coverage_form_var.get().strip(),
        }

    def _read_definitions(self, show_errors=True) -> None:
        excel = Path(self.excel_var.get().strip())
        if not excel.is_file():
            if show_errors:
                messagebox.showwarning("入力なし", "入力Excelを選択してください。", parent=self.root)
            return
        self.status_var.set("Excel定義を読み込み中...")
        self.root.update_idletasks()
        try:
            fields, sheet, header, _columns = read_layout_fields(**self._read_args())
        except Exception as exc:
            self.status_var.set("Excel定義を読めませんでした: %s" % exc)
            if show_errors:
                messagebox.showerror("Excel定義エラー", str(exc), parent=self.root)
            return
        self.loaded_fields = fields
        self.form_fields = OrderedDict()
        for field in fields:
            self.form_fields.setdefault(field.form_id, []).append(field)
        form_ids = list(self.form_fields.keys())
        self.form_box.config(values=form_ids)
        if self.form_var.get() not in self.form_fields:
            self.form_var.set(form_ids[0])
        self._show_selected_form()
        self.status_var.set(
            "読込完了: シート %s / 見出し行 %d / FORM_ID %d件 / 項目 %d件"
            % (sheet, header, len(form_ids), len(fields)))

    def _clear_tree(self) -> None:
        self._finish_cell_edit(save=False)
        for item in self.tree.get_children(""):
            self.tree.delete(item)

    def _show_selected_form(self, _event=None) -> None:
        form_id = self.form_var.get().strip()
        if not self.form_fields:
            return
        if form_id not in self.form_fields:
            messagebox.showwarning("FORM_ID", "Excelに存在するFORM_IDを選択してください。",
                                   parent=self.root)
            return
        self._clear_tree()
        fields = self.form_fields[form_id]
        for field in fields:
            values = (
                "1", field.row_number, field.layout_id, field.field_id,
                field.item_name, field.data_type, field.ime_name,
                "NULL" if field.max_digits is None else field.max_digits,
                field.value, field.attribute_flag, field.coordinates,
            )
            self.tree.insert("", "end", iid=str(field.row_number), values=values)
        self.form_summary_var.set("%s: %d項目（ダブルクリックで編集）" % (form_id, len(fields)))

    def _set_all_includes(self, value: str) -> None:
        self._finish_cell_edit(save=True)
        for item in self.tree.get_children(""):
            values = list(self.tree.item(item, "values"))
            values[0] = value
            self.tree.item(item, values=values)

    def _begin_cell_edit(self, event) -> None:
        region = self.tree.identify_region(event.x, event.y)
        item = self.tree.identify_row(event.y)
        column_token = self.tree.identify_column(event.x)
        if region != "cell" or not item or not column_token:
            return
        index = int(column_token[1:]) - 1
        column_name = TREE_COLUMNS[index]
        if column_name not in EDITABLE_COLUMNS:
            return
        values = list(self.tree.item(item, "values"))
        if column_name == "include":
            values[index] = "0" if str(values[index]) == "1" else "1"
            self.tree.item(item, values=values)
            return
        self._finish_cell_edit(save=True)
        bbox = self.tree.bbox(item, column_token)
        if not bbox:
            return
        x, y, width, height = bbox
        editor = ttk.Entry(self.tree)
        editor.insert(0, values[index])
        editor.select_range(0, "end")
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        self.cell_editor = editor
        self.editor_context = (item, index)
        editor.bind("<Return>", lambda _event: self._finish_cell_edit(save=True))
        editor.bind("<Escape>", lambda _event: self._finish_cell_edit(save=False))
        editor.bind("<FocusOut>", lambda _event: self._finish_cell_edit(save=True))

    def _finish_cell_edit(self, save: bool) -> None:
        if self.cell_editor is None:
            return
        editor = self.cell_editor
        context = self.editor_context
        self.cell_editor = None
        self.editor_context = None
        if save and context is not None:
            item, index = context
            if self.tree.exists(item):
                values = list(self.tree.item(item, "values"))
                values[index] = editor.get()
                self.tree.item(item, values=values)
        editor.destroy()

    def _screen_edits(self):
        self._finish_cell_edit(save=True)
        rows = []  # type: List[int]
        overrides = {}  # type: Dict[int, Dict[str, str]]
        for item in self.tree.get_children(""):
            values = list(self.tree.item(item, "values"))
            if str(values[0]).strip() not in ("1", "ON", "on", "yes", "true"):
                continue
            row_number = int(values[1])
            rows.append(row_number)
            overrides[row_number] = {
                "field_id": values[3],
                "value": values[8],
                "attribute_flag": values[9],
                "coordinates": values[10],
            }
        if not rows:
            raise LayoutTxtError("出力対象項目が0件です。出力列を1にしてください。")
        return rows, overrides

    def _tar_only_changed(self) -> None:
        if self.tar_only_var.get():
            self.create_tar_var.set(True)

    def _validate_paths(self):
        excel = Path(self.excel_var.get().strip())
        output_raw = self.output_var.get().strip()
        if not excel.is_file():
            raise LayoutTxtError("入力Excelを選択してください。")
        if not output_raw:
            raise LayoutTxtError("出力フォルダを選択してください。")
        return excel, Path(output_raw)

    def _generate_selected(self) -> None:
        if not self.form_fields:
            self._read_definitions()
            if not self.form_fields:
                return
        form_id = self.form_var.get().strip()
        try:
            rows, overrides = self._screen_edits()
        except LayoutTxtError as exc:
            messagebox.showerror("設定エラー", str(exc), parent=self.root)
            return
        self._run_generation(
            selected_form_ids=[form_id], selected_rows=rows,
            field_overrides=overrides, error_patterns="none")

    def _generate_all(self) -> None:
        self._run_generation(
            selected_form_ids=None, selected_rows=None, field_overrides=None,
            error_patterns=_value_for_label(
                ERROR_PATTERN_LABELS, self.error_pattern_var.get()))

    def _run_generation(self, selected_form_ids, selected_rows,
                        field_overrides, error_patterns) -> None:
        try:
            excel, output = self._validate_paths()
        except LayoutTxtError as exc:
            messagebox.showwarning("入力なし", str(exc), parent=self.root)
            return
        self.selected_button.config(state="disabled")
        self.all_button.config(state="disabled")
        self.status_var.set("生成中...")
        self.root.update_idletasks()
        try:
            result = generate_layout_txt(
                excel_path=excel, output_dir=output,
                sheet_name=self.sheet_var.get() or None, header_row=self._header_row(),
                form_column=self.form_col_var.get(), layout_column=self.layout_col_var.get(),
                field_column=self.field_col_var.get(), item_column=self.item_col_var.get(),
                data_type_column=self.type_col_var.get(), ime_column=self.ime_col_var.get(),
                max_digits_column=self.max_col_var.get(),
                profile=_value_for_label(PROFILE_LABELS, self.profile_var.get()),
                date_mode=_value_for_label(DATE_MODE_LABELS, self.date_mode_var.get()),
                coverage_form_id=self.coverage_form_var.get().strip(),
                error_patterns=error_patterns,
                filename_template=self.filename_template_var.get().strip(),
                generate_tif=self.generate_tif_var.get(),
                selected_form_ids=selected_form_ids,
                selected_rows=selected_rows, field_overrides=field_overrides,
                create_tar=self.create_tar_var.get() or self.tar_only_var.get(),
                tar_name=self.tar_name_var.get().strip(), tar_only=self.tar_only_var.get(),
                output_format=_value_for_label(FORMAT_LABELS, self.format_var.get()),
                encoding=self.encoding_var.get(), split_by_form=self.split_var.get(),
                overwrite=self.overwrite_var.get())
        except LayoutTxtError as exc:
            self.status_var.set("生成できませんでした: %s" % exc)
            messagebox.showerror("設定エラー", str(exc), parent=self.root)
        except Exception as exc:  # noqa: BLE001 - GUI境界で理由を表示する
            self.status_var.set("生成に失敗しました: %s" % exc)
            messagebox.showerror("生成エラー", str(exc), parent=self.root)
        else:
            _print_result(result)
            tar_text = str(result.tar_file) if result.tar_file is not None else "なし"
            self.status_var.set(
                "完了: FORM_ID %d件 / 項目%d件 / TXT%d件 / TIF%d件 / TAR %s"
                % (result.form_count, result.field_count, len(result.txt_files),
                   len(result.tif_files), tar_text))
            messagebox.showinfo(
                "生成完了",
                "Layoutデータを生成しました。\n\nFORM_ID: %d件\n項目: %d件\n"
                "TXT: %d件\nTIF: %d件\nTAR: %s\n\n%s"
                % (result.form_count, result.field_count, len(result.txt_files),
                   len(result.tif_files), tar_text, output), parent=self.root)
        finally:
            self.selected_button.config(state="normal")
            self.all_button.config(state="normal")


def main(initial_excel: Optional[Path] = None) -> int:
    root = tk.Tk()
    LayoutTxtGui(root, initial_excel=initial_excel)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
