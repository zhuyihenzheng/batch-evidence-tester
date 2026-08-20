# -*- coding: utf-8 -*-
"""Excel -> Layout TXT 生成ツールの単独 GUI。

既存の自動テスト画面とは別プロセスで起動する。Excel が大きい場合や壊れている
場合でも、テスト実行画面を巻き添えにしないため。
"""

import sys
from pathlib import Path
from typing import Optional

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError as exc:  # pragma: no cover - GUI 無し環境
    raise SystemExit("画面を開けません（tkinter が見つかりません）: %s" % exc)

from openpyxl import load_workbook

from .layout_txt import LayoutTxtError, _print_result, generate_layout_txt


PROFILE_LABELS = (
    ("通常値（推奨）", "normal"),
    ("最大桁数ちょうど", "max"),
    ("最大桁数 + 1（異常系）", "over"),
)

FORMAT_LABELS = (
    ("1帳票1行（全値ダブルクォート）", "raw"),
    ("ラベル付き（確認用）", "labeled"),
    ("TSV（1 項目 1 行）", "tsv"),
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


def _value_for_label(pairs, label):
    for item_label, value in pairs:
        if item_label == label:
            return value
    return pairs[0][1]


class LayoutTxtGui(object):
    def __init__(self, root, initial_excel: Optional[Path] = None) -> None:
        self.root = root
        root.title("Layout TXT 生成ツール")
        root.geometry("900x790")
        root.minsize(840, 740)

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
        self.format_var = tk.StringVar(value=FORMAT_LABELS[0][0])
        self.encoding_var = tk.StringVar(value="cp932")
        self.split_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Excel を選択してください。")

        self._build()
        if initial_excel:
            self._load_sheets()

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)

        ttk.Label(outer, text="Excel → Layout TXT", font=("", 16, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(
            outer,
            text="I列: データ型 / J列: IME・入力制限 / K列: 最大桁数 から OCR 値を自動生成します。",
            foreground="#444").grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 14))

        ttk.Label(outer, text="入力 Excel:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(outer, textvariable=self.excel_var).grid(
            row=2, column=1, sticky="we", padx=(10, 6), pady=4)
        ttk.Button(outer, text="参照...", command=self._choose_excel).grid(row=2, column=2, pady=4)

        ttk.Label(outer, text="出力フォルダ:").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(outer, textvariable=self.output_var).grid(
            row=3, column=1, sticky="we", padx=(10, 6), pady=4)
        ttk.Button(outer, text="参照...", command=self._choose_output).grid(row=3, column=2, pady=4)

        ttk.Separator(outer).grid(row=4, column=0, columnspan=3, sticky="we", pady=12)

        settings = ttk.LabelFrame(outer, text="読取設定", padding=10)
        settings.grid(row=5, column=0, columnspan=3, sticky="we")
        for index in range(6):
            settings.columnconfigure(index, weight=1 if index % 2 else 0)

        ttk.Label(settings, text="シート:").grid(row=0, column=0, sticky="w")
        self.sheet_box = ttk.Combobox(settings, textvariable=self.sheet_var, state="readonly", width=24)
        self.sheet_box.grid(row=0, column=1, columnspan=2, sticky="we", padx=(6, 18))
        ttk.Label(settings, text="見出し行:").grid(row=0, column=3, sticky="w")
        ttk.Entry(settings, textvariable=self.header_var, width=10).grid(
            row=0, column=4, sticky="w", padx=(6, 0))

        column_items = (
            ("FormID", self.form_col_var), ("LayoutID", self.layout_col_var),
            ("FieldID", self.field_col_var), ("項目名", self.item_col_var),
            ("データ型", self.type_col_var), ("IME", self.ime_col_var),
            ("最大桁数", self.max_col_var),
        )
        for index, (label, variable) in enumerate(column_items):
            row = 1 + (index // 3)
            pair = index % 3
            col = pair * 2
            ttk.Label(settings, text=label + ":").grid(row=row, column=col, sticky="w", pady=(10, 0))
            ttk.Entry(settings, textvariable=variable, width=12).grid(
                row=row, column=col + 1, sticky="we", padx=(6, 14), pady=(10, 0))

        ttk.Label(
            settings,
            text="列には I のような列記号、見出し名、または auto を指定できます。",
            foreground="#666").grid(row=4, column=0, columnspan=6, sticky="w", pady=(10, 0))

        output = ttk.LabelFrame(outer, text="生成設定", padding=10)
        output.grid(row=6, column=0, columnspan=3, sticky="we", pady=(12, 0))
        output.columnconfigure(1, weight=1)
        output.columnconfigure(3, weight=1)

        ttk.Label(output, text="データ:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(output, textvariable=self.profile_var,
                     values=[label for label, _value in PROFILE_LABELS],
                     state="readonly", width=24).grid(row=0, column=1, sticky="we", padx=(6, 18))
        ttk.Label(output, text="日付形式:").grid(row=0, column=2, sticky="w")
        ttk.Combobox(output, textvariable=self.date_mode_var,
                     values=[label for label, _value in DATE_MODE_LABELS],
                     state="readonly", width=27).grid(
                         row=0, column=3, sticky="we", padx=(6, 0))

        ttk.Label(output, text="TXT 形式:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Combobox(output, textvariable=self.format_var,
                     values=[label for label, _value in FORMAT_LABELS],
                     state="readonly", width=24).grid(
                         row=1, column=1, sticky="we", padx=(6, 18), pady=(10, 0))

        ttk.Label(output, text="文字コード:").grid(row=1, column=2, sticky="w", pady=(10, 0))
        ttk.Combobox(output, textvariable=self.encoding_var,
                     values=("cp932", "utf-8-sig", "utf-8"),
                     state="readonly", width=16).grid(
                         row=1, column=3, sticky="w", padx=(6, 0), pady=(10, 0))
        ttk.Checkbutton(output, text="FormID ごとに分割", variable=self.split_var).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Checkbutton(output, text="既存 TXT / TIF を上書き", variable=self.overwrite_var).grid(
            row=2, column=2, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Label(output, text="全網羅FormID:").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(output, textvariable=self.coverage_form_var, width=12).grid(
            row=3, column=1, sticky="w", padx=(6, 18), pady=(10, 0))
        ttk.Label(output, text="エラーPattern:").grid(row=3, column=2, sticky="w", pady=(10, 0))
        ttk.Combobox(output, textvariable=self.error_pattern_var,
                     values=[label for label, _value in ERROR_PATTERN_LABELS],
                     state="readonly", width=20).grid(
                         row=3, column=3, sticky="we", padx=(6, 0), pady=(10, 0))
        ttk.Label(output, text="ファイル名:").grid(row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(output, textvariable=self.filename_template_var).grid(
            row=4, column=1, sticky="we", padx=(6, 18), pady=(10, 0))
        ttk.Checkbutton(output, text="TXTと同名のTIFを生成",
                        variable=self.generate_tif_var).grid(
                            row=4, column=2, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Label(
            output,
            text="使用可: {form_id} {output_form_id} {pattern} {seq:02d} {source}",
            foreground="#666").grid(
                row=5, column=0, columnspan=4, sticky="w", pady=(8, 0))

        notes = (
            "1 行 = 1 帳票。全値をダブルクォートし、カンマで区切ります。\n"
            "4001は選択したエラーPatternごとにTXT/TIFを1組生成。順序: FormID → 対象有無 → "
            "[ELEMENT_ID → OCR値 → 属性フラグ → 座標] × 項目数"
        )
        ttk.Label(outer, text=notes, foreground="#444").grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(14, 0))

        status = tk.Label(outer, textvariable=self.status_var, anchor="w", justify="left",
                          fg="#444", wraplength=760)
        status.grid(row=8, column=0, columnspan=3, sticky="we", pady=(16, 8))

        buttons = ttk.Frame(outer)
        buttons.grid(row=9, column=0, columnspan=3, sticky="e", pady=(4, 0))
        ttk.Button(buttons, text="閉じる", command=self.root.destroy).pack(side="right", padx=(8, 0))
        self.generate_button = ttk.Button(buttons, text="TXT / TIF を生成", command=self._generate)
        self.generate_button.pack(side="right")

    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root, title="レイアウト定義 Excel を選択",
            filetypes=(("Excel", "*.xlsx *.xlsm"), ("すべて", "*.*")))
        if not path:
            return
        self.excel_var.set(path)
        if not self.output_var.get():
            self.output_var.set(str(Path(path).parent / "layout_txt"))
        self._load_sheets()

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(parent=self.root, title="TXT 出力フォルダを選択")
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
            messagebox.showerror("Excel 読込エラー", str(exc), parent=self.root)
            return
        self.sheet_box.config(values=names)
        self.sheet_var.set(active)
        self.status_var.set("%d シートを読み込みました。列設定を確認して生成してください。" % len(names))

    def _header_row(self):
        raw = self.header_var.get().strip()
        if not raw or raw in ("自動", "auto", "AUTO"):
            return None
        try:
            value = int(raw)
        except ValueError:
            raise LayoutTxtError("見出し行は数値または「自動」にしてください。")
        if value < 1:
            raise LayoutTxtError("見出し行は 1 以上にしてください。")
        return value

    def _generate(self) -> None:
        excel = Path(self.excel_var.get().strip())
        output = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if not excel.is_file():
            messagebox.showwarning("入力なし", "入力 Excel を選択してください。", parent=self.root)
            return
        if output is None:
            messagebox.showwarning("入力なし", "出力フォルダを選択してください。", parent=self.root)
            return

        self.generate_button.config(state="disabled")
        self.status_var.set("生成中...")
        self.root.update_idletasks()
        try:
            result = generate_layout_txt(
                excel_path=excel, output_dir=output,
                sheet_name=self.sheet_var.get() or None, header_row=self._header_row(),
                form_column=self.form_col_var.get(), layout_column=self.layout_col_var.get(),
                field_column=self.field_col_var.get(),
                item_column=self.item_col_var.get(), data_type_column=self.type_col_var.get(),
                ime_column=self.ime_col_var.get(), max_digits_column=self.max_col_var.get(),
                profile=_value_for_label(PROFILE_LABELS, self.profile_var.get()),
                date_mode=_value_for_label(DATE_MODE_LABELS, self.date_mode_var.get()),
                coverage_form_id=self.coverage_form_var.get().strip(),
                error_patterns=_value_for_label(
                    ERROR_PATTERN_LABELS, self.error_pattern_var.get()),
                filename_template=self.filename_template_var.get().strip(),
                generate_tif=self.generate_tif_var.get(),
                output_format=_value_for_label(FORMAT_LABELS, self.format_var.get()),
                encoding=self.encoding_var.get(), split_by_form=self.split_var.get(),
                overwrite=self.overwrite_var.get())
        except LayoutTxtError as exc:
            self.status_var.set("生成できませんでした: %s" % exc)
            messagebox.showerror("設定エラー", str(exc), parent=self.root)
        except Exception as exc:  # noqa: BLE001 - GUI 境界で理由を表示する
            self.status_var.set("生成に失敗しました: %s" % exc)
            messagebox.showerror("生成エラー", str(exc), parent=self.root)
        else:
            _print_result(result)
            self.status_var.set(
                "完了: FormID %d件 / Pattern %d件 / TXT %d件 / TIF %d件\n%s"
                % (result.form_count, result.pattern_count, len(result.txt_files),
                   len(result.tif_files), output))
            messagebox.showinfo(
                "生成完了",
                "Layoutデータを生成しました。\n\nFormID: %d件\nFieldID: %d件\n"
                "Pattern: %d件\nTXT: %d件\nTIF: %d件\n\n%s"
                % (result.form_count, result.field_count, result.pattern_count,
                   len(result.txt_files), len(result.tif_files), output),
                parent=self.root)
        finally:
            self.generate_button.config(state="normal")


def main(initial_excel: Optional[Path] = None) -> int:
    root = tk.Tk()
    LayoutTxtGui(root, initial_excel=initial_excel)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
