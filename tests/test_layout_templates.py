# -*- coding: utf-8 -*-
"""手編集、移動、再出力でテンプレートの値を失わないことを確認する。"""

import csv
import io
import shutil
import sys
import tarfile
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'src'))

from autotest.layout_templates import (TemplateError, save_template, load_template,
                                      list_templates, delete_template,
                                      scan_template_folder, export_template_batch)
from autotest.layout_tar import PackageItem, build_image_tar
from autotest.layout_txt import LayoutField, generate_layout_txt


class TemplateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.path = self.root / 'templates.xlsx'
        self.field = dict(vars(LayoutField('1001', '1', '001', '通院日', '文字列',
                                           '半角英数', 30, '=001\n 値 ', 3)))
        self.field['include'] = '0'
        stream = io.StringIO()
        writer = csv.writer(stream, quoting=csv.QUOTE_ALL, lineterminator='\r\n')
        writer.writerow(['1001', '1', '001', '001\n"値", ', '0', '10,10,60,60'])
        writer.writerow(['1001', '1', '001', '別帳票', '2', '70,10,120,60'])
        self.item = PackageItem('001_', form_id='0099', front_image_bytes=b'front',
                                back_image_bytes=b'back', front_recognition_text=stream.getvalue(),
                                scan_batch_id='000123', back_recognition_result='',
                                extra_fields={'custom': '001'})
        self.settings = {'filename_template_var': '001', 'generate_tif_var': True}

    def tearDown(self):
        self.temp.cleanup()

    def save(self, name='正常', overwrite=False):
        # 旧形式の互換性を検査するため、このヘルパーだけ複数テンプレートを許可する。
        save_template(self.path, name, self.settings, [self.field], [(self.item, '0')],
                      overwrite=overwrite, single_file=False)

    def test_single_file_save_replaces_one_template_and_keeps_other_files(self):
        first = self.root / '正常.xlsx'
        second = self.root / '空値.xlsx'
        save_template(first, '正常', self.settings, [self.field], [(self.item, '1')])
        first_bytes = first.read_bytes()
        self.field['value'] = ''
        save_template(second, '空値', self.settings, [self.field], [(self.item, '0')])
        self.assertEqual(first.read_bytes(), first_bytes)
        self.assertEqual(list_templates(first), ['正常'])
        self.assertEqual(list_templates(second), ['空値'])
        save_template(second, '変更後', self.settings, [self.field], [(self.item, '0')], overwrite=True)
        self.assertEqual(list_templates(second), ['変更後'])
        wb = load_workbook(str(second))
        self.assertEqual(wb['編集項目'].max_row, 2)
        wb.close()
        shutil.rmtree(str(self.root / '正常_assets'))
        self.assertEqual(load_template(second, '変更後')[2][0][0].front_payload(), b'front')

    def test_single_file_refuses_to_overwrite_legacy_multiple_templates(self):
        self.save('正常')
        self.save('空値')
        before = self.path.read_bytes()
        with self.assertRaises(TemplateError):
            save_template(self.path, '新規', self.settings, [self.field], [], overwrite=True)
        self.assertEqual(self.path.read_bytes(), before)

    def test_folder_scan_lists_templates_and_reports_other_excel(self):
        self.save('正常')
        self.save('空値')
        wb = Workbook()
        wb.save(str(self.root / 'other.xlsx'))
        wb.close()
        (self.root / '~$locked.xlsx').write_bytes(b'lock')
        (self.root / 'subdir').mkdir()
        entries, errors = scan_template_folder(self.root)
        self.assertEqual(entries, [(self.path, '正常'), (self.path, '空値')])
        self.assertEqual([p.name for p, _reason in errors], ['other.xlsx'])

    def test_batch_outputs_keep_individual_csv_settings_and_names(self):
        first, second = self.root / 'a.xlsx', self.root / 'b.xlsx'
        for path, value in ((first, '0001'), (second, '0002')):
            self.item.scan_batch_id = value
            save_template(path, path.stem, {
                'package_tar_name_var': 'same', 'manifest_name_var': 'list.csv',
                'manifest_style_var': '固定10列（画像ごと・ヘッダーなし）',
                'manifest_encoding_var': 'utf-8',
            }, [self.field], [(self.item, '1')])
        entries = [(first, 'a'), (second, 'b')]
        result = export_template_batch(entries, self.root / 'out', 'csv')
        for index, (name, value) in enumerate((('a', '0001'), ('b', '0002')), 1):
            with tarfile.open(str(result / ('%03d_%s' % (index, name)) / 'same.tar')) as tar:
                content = tar.extractfile('list.csv').read().decode('utf-8')
                self.assertEqual(next(csv.reader(io.StringIO(content)))[0], value)
                self.assertNotIn(self.item.front_recognition_name, tar.getnames())
        again = export_template_batch(entries, self.root / 'out', 'txt')
        self.assertNotEqual(result, again)
        with tarfile.open(str(next((again / '001_a').glob('*.tar')))) as tar:
            self.assertEqual(tar.extractfile(self.item.front_recognition_name).read().decode('cp932'),
                             self.item.front_recognition_text.replace('\r\n', '\n').replace('\n', '\r\n'))

    def test_failed_batch_does_not_publish_partial_results(self):
        self.save()
        # 最初のテンプレートは成功し、2件目の画像読み取りで失敗する。
        second = self.root / 'broken.xlsx'
        save_template(second, 'broken', self.settings, [self.field], [(self.item, '1')])
        for path in (self.root / 'broken_assets').iterdir():
            path.unlink()
        good = self.root / 'good.xlsx'
        save_template(good, 'good', self.settings, [self.field], [(self.item, '1')])
        output = self.root / 'out'
        output.mkdir()
        (output / 'existing.txt').write_text('keep', encoding='utf-8')
        with self.assertRaises(TemplateError) as error:
            export_template_batch([(good, 'good'), (second, 'broken')], output, 'txt')
        self.assertIn('broken.xlsx', str(error.exception))
        self.assertEqual([p.name for p in output.iterdir()], ['existing.txt'])

    def test_batch_direct_txt_without_output_list(self):
        self.field['include'] = '1'
        save_template(self.path, 'direct', {'generate_tif_var': False, 'create_tar_var': False,
                                          'encoding_var': 'utf-8'}, [self.field], [])
        result = export_template_batch([(self.path, 'direct')], self.root / 'out', 'txt')
        text = (result / '001_templates' / '1001.txt').read_text(encoding='utf-8')
        self.assertEqual(next(csv.reader(io.StringIO(text)))[3], '=001\n 値 ')
        with self.assertRaises(TemplateError):
            export_template_batch([(self.path, 'direct')], self.root / 'out', 'csv')

    def test_roundtrip_preserves_values_and_tar_members(self):
        self.save()
        settings, fields, packages = load_template(self.path, '正常')
        self.assertEqual(settings, self.settings)
        self.assertEqual(fields[0][0].value, '=001\n 値 ')
        self.assertEqual(fields[0][1], '0')
        item, include = packages[0]
        self.assertEqual(include, '0')
        self.assertEqual(item.front_recognition_text, self.item.front_recognition_text)
        self.assertEqual(item.scan_batch_id, '000123')
        self.assertEqual(item.back_recognition_result, '')
        self.assertEqual(item.extra_fields, {'custom': '001'})
        before = build_image_tar([self.item], self.root / 'before', 'same', include_manifest_csv=True)
        after = build_image_tar([item], self.root / 'after', 'same', include_manifest_csv=True)
        self.assertEqual(before.tar_file.read_bytes(), after.tar_file.read_bytes())

    def test_manual_cells_drive_txt_and_blank_clears_attributes(self):
        self.save()
        wb = load_workbook(str(self.path))
        ws = wb['出力OCR']
        ws['H2'] = None
        ws['H3'] = '手動設定'
        wb['出力帳票']['L2'] = '000007'
        wb.save(str(self.path))
        wb.close()
        _, _, packages = load_template(self.path, '正常')
        records = list(csv.reader(io.StringIO(packages[0][0].front_recognition_text)))
        self.assertEqual(records[0][3:], ['', '', ''])
        self.assertEqual(records[1][3], '手動設定')
        self.assertEqual(packages[0][0].scan_batch_id, '000007')

    def test_multiple_templates_overwrite_and_delete_are_isolated(self):
        self.save()
        self.save('空値')
        with self.assertRaises(TemplateError):
            self.save()
        self.field['value'] = '更新'
        self.save(overwrite=True)
        self.assertEqual(load_template(self.path, '正常')[1][0][0].value, '更新')
        self.assertEqual(load_template(self.path, '空値')[1][0][0].value, '=001\n 値 ')
        wb = load_workbook(str(self.path))
        self.assertEqual(wb['編集項目'].max_row, 3)
        wb.close()
        delete_template(self.path, '正常')
        self.assertEqual(list_templates(self.path), ['空値'])

    def test_relative_assets_survive_moving_directory(self):
        self.save()
        moved = self.root / 'moved'
        moved.mkdir()
        shutil.move(str(self.path), str(moved / self.path.name))
        shutil.move(str(self.root / 'templates_assets'), str(moved / 'templates_assets'))
        packages = load_template(moved / self.path.name, '正常')[2]
        self.assertEqual(packages[0][0].front_payload(), b'front')

    def test_invalid_links_and_duplicate_fields_are_rejected(self):
        self.save()
        wb = load_workbook(str(self.path))
        ws = wb['出力OCR']
        ws.append([c.value for c in ws[2]])
        wb.save(str(self.path))
        wb.close()
        with self.assertRaises(TemplateError):
            load_template(self.path, '正常')
        self.save(overwrite=True)
        wb = load_workbook(str(self.path))
        wb['出力帳票']['D2'] = '../outside.tif'
        wb.save(str(self.path))
        wb.close()
        with self.assertRaises(TemplateError):
            load_template(self.path, '正常')

    def test_formula_input_is_rejected_but_literal_equals_is_preserved(self):
        self.save()
        wb = load_workbook(str(self.path))
        wb['編集項目']['I2'] = '=1+1'
        wb.save(str(self.path))
        wb.close()
        with self.assertRaises(TemplateError):
            load_template(self.path, '正常')

    def test_direct_output_does_not_require_original_excel(self):
        self.save()
        field = load_template(self.path, '正常')[1][0][0]
        result = generate_layout_txt(self.root / 'missing.xlsx', self.root / 'out',
                                     prepared_fields=[field], error_patterns='none',
                                     encoding='utf-8', generate_tif=False)
        records = list(csv.reader(io.StringIO(result.txt_files[0].read_text(encoding='utf-8'))))
        self.assertEqual(records[0][3], '=001\n 値 ')

    def test_missing_assets_fail_before_loading(self):
        self.save()
        for asset in (self.root / 'templates_assets').iterdir():
            asset.unlink()
        with self.assertRaises(TemplateError):
            load_template(self.path, '正常')


if __name__ == '__main__':
    unittest.main()
