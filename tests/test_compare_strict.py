# -*- coding: utf-8 -*-
"""比較は前後の空白を外さないこと（偽 OK の回帰テスト）。

以前は比較用の正規化で strip していたため、空白 1 つだけ違う値・行を「一致」と
判定していた。固定長・EDI では末尾の埋め草も、右寄せの数字欄の先頭空白も値の一部で、
そこが欠けた出力を OK にするのは偽 OK そのもの。

吸収してよいのは改行コード（CRLF / LF）だけ。Windows の git（autocrlf）が
取り出した期待値ファイルの改行を書き換えるため。

  python -m unittest tests.test_compare_strict -v
"""

import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest.compare import compare_db_table, compare_text_file  # noqa: E402
from autotest.models import NG, OK, Table  # noqa: E402


def table(columns, *rows):
    return Table(title="T", columns=list(columns), rows=[list(r) for r in rows])


class TestDbValues(unittest.TestCase):
    def test_trailing_space_difference_is_ng_and_explained(self):
        result = compare_db_table("T", table(["ID", "MEMO"], ["1", "ab  "]),
                                  table(["ID", "MEMO"], ["1", "ab "]), keys=["ID"])
        self.assertEqual(result.verdict, NG)
        self.assertIn("前後の空白だけ", result.detail)
        self.assertIn("RTRIM", result.detail)

    def test_leading_space_difference_is_ng(self):
        """右寄せの数字欄（'   123'）を '123' と同一視しない。"""
        result = compare_db_table("T", table(["ID", "AMT"], ["1", "   123"]),
                                  table(["ID", "AMT"], ["1", "123"]), keys=["ID"])
        self.assertEqual(result.verdict, NG)

    def test_char_padded_key_does_not_match_unpadded_key(self):
        result = compare_db_table("T", table(["CODE", "V"], ["A001      ", "x"]),
                                  table(["CODE", "V"], ["A001", "x"]), keys=["CODE"])
        self.assertEqual(result.verdict, NG)
        self.assertIn("前後の空白だけ", result.detail, "埋め草だけの違いだと分かること")

    def test_positional_comparison_is_strict_too(self):
        result = compare_db_table("T", table(["V"], ["x "]), table(["V"], ["x"]), keys=[])
        self.assertEqual(result.verdict, NG)
        self.assertIn("前後の空白だけ", result.detail)

    def test_newline_style_inside_a_value_is_tolerated(self):
        result = compare_db_table("T", table(["ID", "NOTE"], ["1", "a\nb"]),
                                  table(["ID", "NOTE"], ["1", "a\r\nb"]), keys=["ID"])
        self.assertEqual(result.verdict, OK)

    def test_real_difference_has_no_whitespace_note(self):
        result = compare_db_table("T", table(["ID", "V"], ["1", "x"]), table(["ID", "V"], ["1", "y"]), keys=["ID"])
        self.assertEqual(result.verdict, NG)
        self.assertNotIn("空白", result.detail)


class TestFileLines(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="autotest_cmp_"))
        self.addCleanup(shutil.rmtree, str(self.dir), True)

    def compare(self, expected_text, actual_text):
        expected = self.dir / "expected.txt"
        expected.write_bytes(expected_text.encode("utf-8"))
        return compare_text_file("OUT", None, expected, actual_text=actual_text)

    def test_missing_trailing_padding_in_fixed_length_record_is_ng(self):
        result = self.compare("H0001ABC  \nD0002   12\n", "H0001ABC\nD0002   12\n")
        self.assertEqual(result.verdict, NG)
        self.assertIn("前後の空白だけ", result.detail)

    def test_leading_space_in_right_aligned_field_is_ng(self):
        self.assertEqual(self.compare("   12\n", "12\n").verdict, NG)

    def test_crlf_expected_file_still_matches_lf_output(self):
        self.assertEqual(self.compare("a  \r\nb\r\n", "a  \nb\n").verdict, OK)


if __name__ == "__main__":
    unittest.main()
