# -*- coding: utf-8 -*-
"""Excel -> Layout TXT 生成ツールの回帰テスト。"""

import csv
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest.layout_txt import (  # noqa: E402
    LayoutTxtError,
    generate_layout_txt,
    generate_ocr_value,
    read_layout_fields,
)


class LayoutWorkbookCase(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="autotest_layout_txt_"))
        self.book = self.tmp / "layout.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "帳票定義"
        ws.append(["説明行"])
        headers = ["予備A", "FORM_ID", "LAYOUT_ID", "CYOUHYOU_NAME", "LAYOUT_ID",
                   "GROUP_ID", "GROUP_NAME", "ITEM_NAME", "ELEMENT_DATA_TYPE_NAME",
                   "ELEMENT_IME_NAME", "MAX_NUM_DIGITS", "ELEMENT_ID", "CONTROL_ID"]
        ws.append(headers)
        ws.append([None, 1001, 1, "帳票A", 1, 1, "基本", "傷病名1",
                   "文字列", "ひらがな", 100, 1001, 1001])
        ws["C3"].number_format = "00000"
        ws.append([None, 1001, 1, "帳票A", 1, 2, "初診", "退院区分1",
                   "選択肢", "数値のみ", "NULL", 1002, 1002])
        ws["C4"].number_format = "00000"
        ws.append([None, 1001, 1, "帳票A", 1, 2, "初診", "複数選択",
                   "チェックボックス", "全タイプ", "NULL", 1003, 1003])
        ws["C5"].number_format = "00000"
        ws.append([None, 1001, 1, "帳票A", 1, 2, "初診", "生年月日",
                   "日付", "数値のみ", "NULL", 1004, 1004])
        ws["C6"].number_format = "00000"
        ws.append([None, 1001, 1, "帳票A", 1, 2, "初診", "入院期間1",
                   "日付 (From)", "数値のみ", "NULL", 1005, 1005])
        ws["C7"].number_format = "00000"
        ws.append([None, 1001, 1, "帳票A", 1, 2, "初診", "退院期間1",
                   "日付 (To)", "数値のみ", "NULL", 1006, 1006])
        ws["C8"].number_format = "00000"
        ws.append([None, 1001, 1, "帳票A", 1, 2, "初診", "処置日",
                   "カレンダー", "数値のみ", "NULL", 1007, 1007])
        ws["C9"].number_format = "00000"
        ws.append([None, 4001, 2, "全網羅帳票", 2, 1, "基本", "網羅日付1",
                   "日付", "数値のみ", "NULL", 2001, 2001])
        ws["C10"].number_format = "00000"
        ws.append([None, 4001, 2, "全網羅帳票", 2, 1, "基本", "網羅日付2",
                   "日付", "数値のみ", "NULL", 2002, 2002])
        ws["C11"].number_format = "00000"
        ws.append([None, 4001, 2, "全網羅帳票", 2, 1, "基本", "網羅日付3",
                   "日付", "数値のみ", "NULL", 2003, 2003])
        ws["C12"].number_format = "00000"
        ws.append([None, 4001, 2, "全網羅帳票", 2, 1, "基本", "網羅日付4",
                   "日付", "数値のみ", "NULL", 2004, 2004])
        ws["C13"].number_format = "00000"
        wb.save(str(self.book))

    def tearDown(self):
        shutil.rmtree(str(self.tmp), ignore_errors=True)


class TestReadLayoutFields(LayoutWorkbookCase):
    def test_auto_detects_header_and_identifier_columns(self):
        fields, sheet, header, columns = read_layout_fields(self.book)
        self.assertEqual(sheet, "帳票定義")
        self.assertEqual(header, 2)
        self.assertEqual(columns["form_id"], 2)
        self.assertEqual(columns["layout_id"], 3)
        self.assertEqual(columns["field_id"], 12)
        self.assertEqual(columns["data_type"], 9)
        self.assertEqual(columns["ime"], 10)
        self.assertEqual(columns["max_digits"], 11)
        self.assertEqual([field.form_id for field in fields],
                         ["1001"] * 7 + ["4001"] * 4)
        self.assertEqual([field.layout_id for field in fields],
                         ["00001"] * 7 + ["00002"] * 4)
        self.assertEqual([field.field_id for field in fields],
                         ["1001", "1002", "1003", "1004", "1005", "1006", "1007",
                          "2001", "2002", "2003", "2004"])
        self.assertEqual([field.value for field in fields],
                         ["傷病名1", "1", "1.2", "5/8/6/1", "5/8/6/1",
                          "5/8/6/1", "5/8/6/1", "5/8/6/1", "/2026/6/1",
                          "5/2026/6/1", "5/8/6/1|5/8/6/2"])
        self.assertEqual([field.attribute_flag for field in fields],
                         ["0"] * 11)

    def test_column_can_be_selected_by_header_or_letter(self):
        fields, _sheet, _header, columns = read_layout_fields(
            self.book, form_column="B", layout_column="C", field_column="ELEMENT_ID",
            data_type_column="ELEMENT_DATA_TYPE_NAME", ime_column="J",
            max_digits_column="MAX_NUM_DIGITS")
        self.assertEqual(len(fields), 11)
        self.assertEqual(columns["field_id"], 12)

    def test_missing_sheet_is_reported_with_candidates(self):
        with self.assertRaises(LayoutTxtError) as ctx:
            read_layout_fields(self.book, sheet_name="存在しない")
        self.assertIn("帳票定義", str(ctx.exception))


class TestValueGeneration(unittest.TestCase):
    def test_supported_ime_rules(self):
        cases = (
            ("文字列", "ひらがな", 5, "あいうえお"),
            ("文字列", "半角カタカナ", 5, "ｱｲｳｴｵ"),
            ("文字列", "半角英数", 6, "Abc123"),
            ("文字列", "数値のみ", 4, "1234"),
            ("文字列", "小数点数値", 6, "123.45"),
            ("文字列", "全タイプ", 4, "あA1ｱ"),
            ("チェックボックス", "NULL", None, "1.2"),
            ("選択肢", "数値のみ", None, "1"),
            ("カレンダー", "数値のみ", None, "5/8/6/1"),
        )
        for data_type, ime, digits, expected in cases:
            self.assertEqual(generate_ocr_value(data_type, ime, digits), expected)

    def test_max_and_over_profiles_hit_boundary(self):
        self.assertEqual(generate_ocr_value("文字列", "ひらがな", 7, "max"), "あいうえおあい")
        self.assertEqual(len(generate_ocr_value("文字列", "ひらがな", 7, "over")), 8)
        self.assertEqual(generate_ocr_value("文字列", "小数点数値", 5, "max"), "1.234")
        self.assertEqual(len(generate_ocr_value("文字列", "数値のみ", 5, "over")), 6)

    def test_string_field_uses_item_name_as_normal_value(self):
        self.assertEqual(
            generate_ocr_value("文字列", "ひらがな", 100, item_name="傷病名1"),
            "傷病名1")
        self.assertEqual(
            generate_ocr_value("文字列", "ひらがな", 3, item_name="医療機関名称1"),
            "医療機")
        self.assertEqual(
            generate_ocr_value("文字列", "ひらがな", 7, "max", item_name="傷病名1"),
            "傷病名1傷病名")

    def test_four_date_notations(self):
        expected = ("5/8/6/1", "/2026/6/1", "5/2026/6/1", "5/8/6/1|5/8/6/2")
        for index, value in enumerate(expected):
            self.assertEqual(
                generate_ocr_value("日付", "数値のみ", None,
                                   date_mode="cycle", date_index=index),
                value)
        self.assertEqual(
            generate_ocr_value("日付", "数値のみ", None, date_mode="multiple"),
            "5/8/6/1|5/8/6/2")


class TestGenerateLayoutTxt(LayoutWorkbookCase):
    def test_splits_one_raw_file_per_form_with_crlf_and_cp932(self):
        out = self.tmp / "out"
        result = generate_layout_txt(self.book, out, generate_tif=False)
        self.assertEqual(result.form_count, 2)
        self.assertEqual(result.field_count, 11)
        self.assertEqual(result.pattern_count, 25)
        self.assertEqual(len(result.txt_files), 25)
        self.assertEqual(result.txt_files[0].name, "1001.txt")
        self.assertEqual(result.txt_files[1].name, "4001_01_normal.txt")
        self.assertEqual(result.txt_files[-1].name, "4001_24_form_id_empty.txt")
        self.assertEqual(result.tif_files, [])

        text = (out / "1001.txt").read_bytes().decode("cp932")
        self.assertEqual(
            text,
            '"1001","1","1001","傷病名1","0","0,0,0,0",'
            '"1002","1","0","0,0,0,0",'
            '"1003","1.2","0","0,0,0,0",'
            '"1004","5/8/6/1","0","0,0,0,0",'
            '"1005","5/8/6/1","0","0,0,0,0",'
            '"1006","5/8/6/1","0","0,0,0,0",'
            '"1007","5/8/6/1","0","0,0,0,0"\r\n')

        coverage = (out / "4001_01_normal.txt").read_bytes().decode("cp932")
        self.assertEqual(
            coverage,
            '"4001","1","2001","5/8/6/1","0","0,0,0,0",'
            '"2002","/2026/6/1","0","0,0,0,0",'
            '"2003","5/2026/6/1","0","0,0,0,0",'
            '"2004","5/8/6/1|5/8/6/2","0","0,0,0,0"\r\n')

    def test_count_and_element_id_patterns_change_actual_structure(self):
        out = self.tmp / "patterns"
        generate_layout_txt(self.book, out, generate_tif=False)

        zero = next(csv.reader([
            (out / "4001_02_count_zero.txt").read_text(encoding="cp932").strip()]))
        missing = next(csv.reader([
            (out / "4001_03_count_missing.txt").read_text(encoding="cp932").strip()]))
        extra = next(csv.reader([
            (out / "4001_04_count_extra.txt").read_text(encoding="cp932").strip()]))
        unknown = next(csv.reader([
            (out / "4001_05_element_id_unknown.txt").read_text(encoding="cp932").strip()]))
        duplicate = next(csv.reader([
            (out / "4001_06_element_id_duplicate.txt").read_text(encoding="cp932").strip()]))

        self.assertEqual(len(zero), 2)
        self.assertEqual((len(missing) - 2) // 4, 3)
        self.assertEqual((len(extra) - 2) // 4, 5)
        self.assertEqual(missing[-2], "1")
        self.assertEqual(extra[-2], "1")
        self.assertNotIn(unknown[2], ("2001", "2002", "2003", "2004"))
        self.assertEqual(duplicate[2], duplicate[6])

    def test_mixed_attribute_uses_quoted_comma_as_one_value(self):
        out = self.tmp / "mixed"
        generate_layout_txt(self.book, out, generate_tif=False)
        raw = (out / "4001_08_attribute_mixed_1_2.txt").read_text(encoding="cp932")
        values = next(csv.reader([raw.strip()]))
        self.assertEqual(values[4::4], ["0", "1", "2", "1,2"])
        self.assertIn('"1,2"', raw)

    def test_filename_template_and_matching_tif(self):
        from PIL import Image

        out = self.tmp / "named"
        result = generate_layout_txt(
            self.book, out, error_patterns="none",
            filename_template="TEST_{form_id}_{seq:02d}")
        self.assertEqual(
            [path.name for path in result.txt_files],
            ["TEST_1001_01.txt", "TEST_4001_01.txt"])
        self.assertEqual(
            [path.name for path in result.tif_files],
            ["TEST_1001_01.tif", "TEST_4001_01.tif"])
        for txt_path, tif_path in zip(result.txt_files, result.tif_files):
            self.assertEqual(txt_path.stem, tif_path.stem)
            with Image.open(str(tif_path)) as image:
                self.assertEqual(image.format, "TIFF")
                self.assertEqual(image.size, (1654, 2339))

    def test_single_labeled_file_contains_each_form(self):
        out = self.tmp / "single"
        result = generate_layout_txt(
            self.book, out, split_by_form=False, output_format="labeled", encoding="utf-8",
            error_patterns="none", generate_tif=False)
        self.assertEqual(len(result.files), 1)
        text = result.files[0].read_text(encoding="utf-8")
        self.assertIn("FormID=1001", text)
        self.assertIn('"FieldID=1001"', text)
        self.assertIn("FormID=4001", text)
        self.assertIn('"OCRText=5/8/6/1|5/8/6/2"', text)
        self.assertEqual(len(text.splitlines()), 2, "1 行が 1 帳票であること")

    def test_existing_file_is_not_overwritten_without_permission(self):
        out = self.tmp / "existing"
        out.mkdir()
        target = out / "1001.txt"
        target.write_text("保護", encoding="utf-8")
        with self.assertRaises(LayoutTxtError) as ctx:
            generate_layout_txt(self.book, out, generate_tif=False)
        self.assertIn("上書きしません", str(ctx.exception))
        self.assertEqual(target.read_text(encoding="utf-8"), "保護")
        self.assertFalse((out / "4001_01_normal.txt").exists(),
                         "失敗時に一部だけ生成してはいけない")

    def test_overwrite_replaces_existing_file(self):
        out = self.tmp / "overwrite"
        out.mkdir()
        target = out / "1001.txt"
        target.write_text("old", encoding="utf-8")
        generate_layout_txt(self.book, out, overwrite=True, generate_tif=False)
        self.assertNotEqual(target.read_text(encoding="cp932"), "old")

    def test_tsv_contains_one_row_per_field(self):
        out = self.tmp / "tsv"
        result = generate_layout_txt(
            self.book, out, output_format="tsv", encoding="utf-8", generate_tif=False)
        rows = result.files[0].read_text(encoding="utf-8").splitlines()
        self.assertEqual(rows[0], "FormID\tTargetPresence\tFieldID\tOCRText\tAttributeFlag\tCoordinates")
        self.assertEqual(len(rows), 8)


if __name__ == "__main__":
    unittest.main()
