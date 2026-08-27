# -*- coding: utf-8 -*-

import csv
import io
import shutil
import sys
import tarfile
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest.layout_tar import (  # noqa: E402
    LayoutTarError,
    PackageItem,
    base_name_from_front,
    build_image_tar,
    format_extra_fields,
    format_package_tar_name,
    matching_back_image,
    matching_recognition_file,
    parse_extra_fields,
    parse_manifest_columns,
)


class TestLayoutImageTar(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="layout_tar_"))

    def tearDown(self):
        shutil.rmtree(str(self.tmp), ignore_errors=True)

    def _members(self, path):
        with tarfile.open(str(path), "r") as archive:
            return archive.getnames()

    def test_txt_and_csv_packages_use_separate_file_setting_names(self):
        self.assertEqual(
            format_package_tar_name(
                "TXT_{source}_{form_id}", "CSV_{source}_{form_id}",
                False, "layout", ["1001"]),
            "TXT_layout_1001")
        self.assertEqual(
            format_package_tar_name(
                "TXT_{source}_{form_id}", "CSV_{source}_{form_id}",
                True, "layout", ["1001"]),
            "CSV_layout_1001")
        self.assertEqual(
            format_package_tar_name("", "", False, "layout", ["1001"]),
            "layout_layout_data")

    def test_front_back_images_use_f_r_and_result_txt_is_separate(self):
        external_front = self.tmp / "scanF.png"
        external_back = self.tmp / "scanR.png"
        external_front.write_bytes(b"FRONT")
        external_back.write_bytes(b"BACK")
        items = [
            PackageItem(
                base_name="FORM1001", form_id="1001",
                front_recognition_text='"1001","1","1001","VALUE","0","0,0,0,1"\r\n',
                back_recognition_result="1",
                front_image_bytes=b"TIF-F", back_image_bytes=b"TIF-R",
                front_extension=".tif", source_label="FORM_ID 1001"),
            PackageItem(
                base_name="scan", form_id="2001",
                front_recognition_text="FRONT_RESULT",
                back_recognition_result="1,2",
                front_image_path=external_front, back_image_path=external_back,
                front_extension=".png"),
        ]

        result = build_image_tar(items, self.tmp, "CUSTOM_PACKAGE")

        self.assertEqual(result.tar_file.name, "CUSTOM_PACKAGE.tar")
        self.assertEqual(result.view_folder, self.tmp / "CUSTOM_PACKAGE")
        self.assertEqual(result.archive_members, [
            "FORM1001F.tif", "FORM1001F.txt", "FORM1001R.tif", "FORM1001R.txt",
            "scanF.png", "scanF.txt", "scanR.png", "scanR.txt",
        ])
        self.assertEqual(
            sorted(path.name for path in result.view_folder.iterdir()),
            sorted(result.archive_members))
        with tarfile.open(str(result.tar_file), "r") as archive:
            for name in result.archive_members:
                self.assertEqual(
                    (result.view_folder / name).read_bytes(),
                    archive.extractfile(name).read())
            self.assertEqual(archive.extractfile("FORM1001F.tif").read(), b"TIF-F")
            self.assertEqual(archive.extractfile("FORM1001R.tif").read(), b"TIF-R")
            self.assertEqual(
                archive.extractfile("FORM1001F.txt").read(),
                b'"1001","1","1001","VALUE","0","0,0,0,1"\r\n')
            self.assertEqual(archive.extractfile("FORM1001R.txt").read(), b'"1"\r\n')
            self.assertEqual(archive.extractfile("scanF.txt").read(), b"FRONT_RESULT\r\n")
            self.assertEqual(archive.extractfile("scanR.txt").read(), b'"1,2"\r\n')

    def test_images_and_custom_manifest_csv_are_packaged(self):
        items = [
            PackageItem(
                base_name="A", form_id="1001",
                front_image_bytes=b"A", related_file="REF_A.dat",
                extra_fields={"case_type": "normal"}),
            PackageItem(
                base_name="B", form_id="2001", back_recognition_result="9",
                front_image_bytes=b"B-F", back_image_bytes=b"B-R",
                related_file="REF_B.dat",
                extra_fields={"case_type": "error"}),
        ]
        columns = parse_manifest_columns(
            "正面=front_image_file,背面=back_image_file,ID=form_id,"
            "背面結果=back_recognition_result,関連=related_file,区分=case_type")

        result = build_image_tar(
            items, self.tmp, "LIST_PACKAGE.tar",
            include_recognition_txt=False,
            include_manifest_csv=True,
            manifest_name="custom_list.csv",
            manifest_columns=columns,
            csv_encoding="utf-8-sig")

        self.assertEqual(
            result.archive_members,
            ["AF.tif", "BF.tif", "BR.tif", "custom_list.csv"])
        with tarfile.open(str(result.tar_file), "r") as archive:
            raw = archive.extractfile("custom_list.csv").read().decode("utf-8-sig")
        rows = list(csv.reader(io.StringIO(raw)))
        self.assertEqual(rows, [
            ["正面", "背面", "ID", "背面結果", "関連", "区分"],
            ["AF.tif", "", "1001", "", "REF_A.dat", "normal"],
            ["BF.tif", "BR.tif", "2001", "9", "REF_B.dat", "error"],
        ])

    def test_fixed_image_list_csv_has_ten_quoted_columns_without_header(self):
        item = PackageItem(
            base_name="2020123100001_001_", form_id="CB147",
            front_image_bytes=b"F", back_image_bytes=b"R",
            front_extension=".TIF", scan_batch_id="2020123100001",
            image_sequence="001", arrival_date="20210611")

        result = build_image_tar(
            [item], self.tmp, "image_list",
            include_recognition_txt=False,
            include_manifest_csv=True,
            manifest_name="CLM_PUNCH_IMAGE_INFO_01.csv",
            manifest_style="image_list",
            csv_encoding="cp932")

        self.assertEqual(result.archive_members, [
            "2020123100001_001_F.TIF",
            "2020123100001_001_R.TIF",
            "CLM_PUNCH_IMAGE_INFO_01.csv",
        ])
        self.assertFalse(any(name.lower().endswith(".txt")
                             for name in result.archive_members))
        self.assertEqual(
            sorted(path.name for path in result.view_folder.iterdir()),
            sorted(result.archive_members))
        with tarfile.open(str(result.tar_file), "r") as archive:
            raw = archive.extractfile(
                "CLM_PUNCH_IMAGE_INFO_01.csv").read().decode("cp932")
        self.assertEqual(
            raw,
            '"2020123100001","001","2020123100001_001_F.TIF",'
            '"20210611","CB147","","","","",""\r\n'
            '"2020123100001","001","2020123100001_001_R.TIF",'
            '"20210611","CB147","","","","",""\r\n')

    def test_fixed_image_list_csv_uses_auto_sequence_and_rejects_newlines(self):
        first = PackageItem(base_name="A_", front_image_bytes=b"A")
        second = PackageItem(base_name="B_", front_image_bytes=b"B")
        result = build_image_tar(
            [first, second], self.tmp, "auto_sequence",
            include_recognition_txt=False, include_manifest_csv=True,
            manifest_style="image_list", csv_encoding="utf-8")
        with tarfile.open(str(result.tar_file), "r") as archive:
            raw = archive.extractfile("file_list.csv").read().decode("utf-8")
        rows = list(csv.reader(io.StringIO(raw)))
        self.assertEqual(rows[0][1], "001")
        self.assertEqual(rows[1][1], "002")
        self.assertEqual(len(rows[0]), 10)

        invalid = PackageItem(
            base_name="C_", front_image_bytes=b"C",
            application_number="line1\nline2")
        with self.assertRaises(LayoutTarError):
            build_image_tar(
                [invalid], self.tmp, "invalid_line",
                include_recognition_txt=False, include_manifest_csv=True,
                manifest_style="image_list")

    def test_front_only_keeps_form_txt_and_does_not_create_r_files(self):
        item = PackageItem(
            base_name="ONLY", form_id="1001",
            front_image_bytes=b"F",
            front_recognition_text='"1001","1","1001","VALUE","0","0,0,0,1"')

        result = build_image_tar([item], self.tmp, "front_only")

        self.assertEqual(result.archive_members, ["ONLYF.tif", "ONLYF.txt"])
        with tarfile.open(str(result.tar_file), "r") as archive:
            self.assertEqual(
                archive.extractfile("ONLYF.txt").read(),
                b'"1001","1","1001","VALUE","0","0,0,0,1"\r\n')
        self.assertFalse(any(name.startswith("ONLYR") for name in result.archive_members))

    def test_manifest_can_list_side_specific_txt_names(self):
        item = PackageItem(
            base_name="PAIR", form_id="1001",
            front_image_bytes=b"F", back_image_bytes=b"R",
            front_recognition_text="FORM_RESULT", back_recognition_result="1")
        columns = parse_manifest_columns(
            "正面TXT=front_recognition_file,背面TXT=back_recognition_file,"
            "背面値=back_recognition_result")

        result = build_image_tar(
            [item], self.tmp, "with_txt_list",
            include_recognition_txt=True, include_manifest_csv=True,
            manifest_columns=columns, csv_encoding="utf-8")

        with tarfile.open(str(result.tar_file), "r") as archive:
            raw = archive.extractfile("file_list.csv").read().decode("utf-8")
        self.assertEqual(list(csv.reader(io.StringIO(raw))), [
            ["正面TXT", "背面TXT", "背面値"],
            ["PAIRF.txt", "PAIRR.txt", "1"],
        ])

    def test_duplicate_member_name_is_rejected_without_writing_tar(self):
        items = [
            PackageItem(base_name="Same", front_image_bytes=b"1",
                        front_recognition_text="A"),
            PackageItem(base_name="same", front_image_bytes=b"2",
                        front_recognition_text="B"),
        ]
        with self.assertRaises(LayoutTarError) as ctx:
            build_image_tar(items, self.tmp, "duplicate")
        self.assertIn("重複", str(ctx.exception))
        self.assertFalse((self.tmp / "duplicate.tar").exists())

    def test_missing_external_image_is_rejected_without_writing_tar(self):
        item = PackageItem(
            base_name="missing", front_image_path=self.tmp / "none.tif")
        with self.assertRaises(LayoutTarError):
            build_image_tar([item], self.tmp, "missing_package")
        self.assertFalse((self.tmp / "missing_package.tar").exists())

    def test_existing_tar_is_protected_until_overwrite_is_explicit(self):
        target = self.tmp / "protected.tar"
        target.write_bytes(b"old")
        item = PackageItem(
            base_name="A", front_image_bytes=b"A", front_recognition_text="A")
        with self.assertRaises(LayoutTarError):
            build_image_tar([item], self.tmp, "protected")
        self.assertEqual(target.read_bytes(), b"old")

        build_image_tar([item], self.tmp, "protected", overwrite=True)
        self.assertNotEqual(target.read_bytes(), b"old")

    def test_existing_view_folder_is_protected_and_replaced_with_overwrite(self):
        folder = self.tmp / "folder_protected"
        folder.mkdir()
        marker = folder / "old.dat"
        marker.write_bytes(b"old")
        item = PackageItem(
            base_name="A", front_image_bytes=b"A", front_recognition_text="TXT")

        with self.assertRaises(LayoutTarError):
            build_image_tar([item], self.tmp, "folder_protected")
        self.assertFalse((self.tmp / "folder_protected.tar").exists())
        self.assertEqual(marker.read_bytes(), b"old")

        result = build_image_tar(
            [item], self.tmp, "folder_protected", overwrite=True)
        self.assertFalse(marker.exists())
        self.assertEqual(
            sorted(path.name for path in result.view_folder.iterdir()),
            ["AF.tif", "AF.txt"])
        self.assertFalse(any(
            path.name.startswith(".folder_protected.")
            for path in self.tmp.iterdir()))

    def test_view_folder_can_be_disabled_for_library_callers(self):
        item = PackageItem(
            base_name="A", front_image_bytes=b"A", front_recognition_text="TXT")
        result = build_image_tar(
            [item], self.tmp, "tar_only", create_view_folder=False)
        self.assertIsNone(result.view_folder)
        self.assertTrue(result.tar_file.is_file())
        self.assertFalse((self.tmp / "tar_only").exists())

    def test_extra_fields_round_trip_and_invalid_input(self):
        values = parse_extra_fields("case_type=normal; relation=ABC_R.txt")
        self.assertEqual(values, {"case_type": "normal", "relation": "ABC_R.txt"})
        self.assertEqual(
            parse_extra_fields(format_extra_fields(values)), values)
        with self.assertRaises(LayoutTarError):
            parse_extra_fields("missing_separator")

    def test_manifest_columns_can_have_custom_headers(self):
        self.assertEqual(
            parse_manifest_columns("画像=image_file,form_id,任意=custom"),
            [("画像", "image_file"), ("form_id", "form_id"), ("任意", "custom")])
        with self.assertRaises(LayoutTarError):
            parse_manifest_columns("A=image_file,A=form_id")

    def test_front_file_can_find_matching_back_and_result(self):
        front = self.tmp / "DOC_01F.tif"
        back = self.tmp / "DOC_01R.tif"
        front_result = self.tmp / "DOC_01F.txt"
        legacy_result = self.tmp / "DOC_01.txt"
        back_result = self.tmp / "DOC_01R.txt"
        front.write_bytes(b"F")
        back.write_bytes(b"R")
        front_result.write_text("FRONT", encoding="cp932")
        legacy_result.write_text("LEGACY", encoding="cp932")
        back_result.write_text("1", encoding="cp932")

        self.assertEqual(base_name_from_front(front), "DOC_01")
        self.assertEqual(matching_back_image(front), back)
        self.assertEqual(matching_recognition_file(front), front_result)
        self.assertEqual(matching_recognition_file(back), back_result)

    def test_txt_package_requires_front_form_result(self):
        item = PackageItem(base_name="A", front_image_bytes=b"A")
        with self.assertRaises(LayoutTarError) as ctx:
            build_image_tar([item], self.tmp, "missing_front_txt")
        self.assertIn("正面TXT", str(ctx.exception))
        self.assertFalse((self.tmp / "missing_front_txt.tar").exists())


if __name__ == "__main__":
    unittest.main()
