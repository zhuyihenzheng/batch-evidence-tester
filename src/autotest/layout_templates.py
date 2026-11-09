# -*- coding: utf-8 -*-
"""手編集できるExcelテンプレートと相対パスの画像資材。"""

import csv
import hashlib
import io
import os
import tempfile
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .layout_txt import LayoutField, _atomic_write
from .layout_tar import PackageItem, _form_records, format_extra_fields, parse_extra_fields


class TemplateError(ValueError):
    pass


FIELD_KEYS = (
    'form_id', 'layout_id', 'field_id', 'item_name', 'data_type', 'ime_name',
    'max_digits', 'value', 'row_number', 'attribute_flag', 'coordinates',
    'input_attribute', 'input_rule', 'notes', 'output_example',
    'occurrence_index', 'occurrence_count', 'source_element_id', 'include')
FIELD_LABELS = (
    'FORM_ID', 'LAYOUT_ID', 'FIELD_ID', '項目名', 'データ型', 'IME', '最大桁数',
    'OCR値', '元Excel行', '属性', '座標', '入力属性', '入力規則', '補足', '出力例',
    '展開番号', '展開件数', '元ELEMENT_ID', '出力')
PACKAGE_KEYS = (
    'base_name', 'form_id', 'back_recognition_result', 'related_file',
    'scan_batch_id', 'image_sequence', 'arrival_date', 'application_number',
    'reception_number', 'format_id', 'delivery_date', 'delivery_shot')
PACKAGE_LABELS = (
    '画像基礎名', 'CSV位置5', '背面OCR値', '関連ファイル', 'CSV位置1', 'CSV位置2',
    'CSV位置4', 'CSV位置6', 'CSV位置7', 'CSV位置8', 'CSV位置9', 'CSV位置10')
SCHEMAS = OrderedDict((
    ('テンプレート設定', ('テンプレート名', '設定キー', '説明', '型', '値')),
    ('編集項目', ('テンプレート名',) + FIELD_LABELS),
    ('出力帳票', ('テンプレート名', '画像ID', '出力', '正面画像', '背面画像',
              'TXTあり', 'CSV追加項目') + PACKAGE_LABELS),
    ('出力OCR', ('テンプレート名', '画像ID', '帳票順', 'FORM_ID', '対象有無',
             '項目順', 'FIELD_ID', 'OCR値', '属性', '座標')),
))


def _text(cell):
    if cell.data_type == 'f':
        raise TemplateError('数式は使用できません。値を入力してください: %s' % cell.coordinate)
    return '' if cell.value is None else str(cell.value)


def _tables(wb):
    tables = {}
    for title, headers in SCHEMAS.items():
        if title not in wb.sheetnames:
            raise TemplateError('テンプレートSheetがありません: %s' % title)
        ws = wb[title]
        actual = [_text(c) for c in ws[1]]
        if actual != list(headers):
            raise TemplateError('見出しを確認してください: %s' % title)
        rows = []
        for cells in ws.iter_rows(min_row=2, max_col=len(headers)):
            row = [_text(c) for c in cells]
            if not any(row):
                continue
            if not row[0]:
                raise TemplateError('テンプレート名が空です: %s / %s' % (title, cells[0].row))
            rows.append(row)
        tables[title] = rows
    return tables


def _open(path):
    path = Path(path)
    if path.suffix.lower() != '.xlsx':
        raise TemplateError('テンプレートは .xlsx を指定してください。')
    return load_workbook(str(path), data_only=False)


def list_templates(path):
    if not Path(path).exists():
        return []
    wb = _open(path)
    try:
        tables = _tables(wb)
        return list(OrderedDict((row[0], None) for row in tables['テンプレート設定']))
    finally:
        wb.close()


def _cell(ws, row, column, value):
    text = str(value)
    if len(text) > 32767:
        raise TemplateError('セルの文字数が32767文字を超えています。')
    cell = ws.cell(row=row, column=column)
    cell.value = text
    cell.data_type = 's'
    cell.number_format = '@'


def _write(path, tables):
    if Path(path).exists():
        wb = _open(path)
        for title in SCHEMAS:
            wb.remove(wb[title])
    else:
        wb = Workbook()
        wb.remove(wb.active)
    temp = None
    try:
        for title, headers in SCHEMAS.items():
            ws = wb.create_sheet(title)
            for row_index, values in enumerate([headers] + tables[title], 1):
                for column, value in enumerate(values, 1):
                    _cell(ws, row_index, column, value)
            ws.freeze_panes = 'B2'
            ws.auto_filter.ref = ws.dimensions
            for column, header in enumerate(headers, 1):
                ws.column_dimensions[get_column_letter(column)].width = (
                    42 if header in ('値', 'OCR値', '正面画像', '背面画像') else 20)
                cell = ws.cell(row=1, column=column)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='315A7D')
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        handle, temp = tempfile.mkstemp(suffix='.xlsx', dir=str(path.parent))
        os.close(handle)
        wb.save(temp)
        os.replace(temp, str(path))
        temp = None
    finally:
        wb.close()
        if temp and os.path.exists(temp):
            os.unlink(temp)


def _asset(path, payload, extension):
    # 内容アドレスにより、別テンプレートや保存失敗で既存画像を上書きしない。
    relative = Path(path.stem + '_assets') / (hashlib.sha256(payload).hexdigest() + extension)
    target = path.parent / relative
    if not target.exists():
        _atomic_write(target, payload)
    elif target.read_bytes() != payload:
        raise TemplateError('画像資材が変更されています: %s' % target)
    return relative.as_posix()


def save_template(path, name, settings, fields, packages, descriptions=None, overwrite=False):
    path = Path(path)
    name = name.strip()
    if not name:
        raise TemplateError('テンプレート名を入力してください。')
    if path.suffix.lower() != '.xlsx':
        raise TemplateError('テンプレートは .xlsx を指定してください。')
    tables = {title: [] for title in SCHEMAS}
    if path.exists():
        wb = _open(path)
        try:
            tables = _tables(wb)
        finally:
            wb.close()
    exists = any(row[0] == name for rows in tables.values() for row in rows)
    if exists and not overwrite:
        raise TemplateError('同名テンプレートがあります: %s' % name)
    tables = {title: [r for r in rows if r[0] != name] for title, rows in tables.items()}
    settings = dict(settings)
    settings['template_version'] = '1'
    for key, value in sorted(settings.items()):
        kind = 'bool' if isinstance(value, bool) else 'str'
        text = ('1' if value else '0') if kind == 'bool' else str(value)
        tables['テンプレート設定'].append([name, key, (descriptions or {}).get(key, key), kind, text])
    for field in fields:
        tables['編集項目'].append([name] + [
            '' if field.get(key) is None else str(field.get(key, '')) for key in FIELD_KEYS])
    for index, (item, include) in enumerate(packages, 1):
        image_id = str(index)
        front = _asset(path, item.front_payload(), item.front_extension)
        back = _asset(path, item.back_payload(), item.back_extension) if item.has_back_image else ''
        tables['出力帳票'].append(
            [name, image_id, include, front, back, '1' if item.has_front_recognition else '0',
             format_extra_fields(item.extra_fields)] + [getattr(item, key) for key in PACKAGE_KEYS])
        if item.has_front_recognition:
            records = _form_records(item.front_recognition_text)
            if not records:
                raise TemplateError('正面TXTが空です: %s' % item.front_recognition_name)
            for record_index, record in enumerate(records, 1):
                chunks = [record[i:i + 4] for i in range(2, len(record), 4)]
                if not chunks:
                    tables['出力OCR'].append([name, image_id, str(record_index)] + record[:2] + ['0', '', '', '', ''])
                for field_index, values in enumerate(chunks, 1):
                    tables['出力OCR'].append([name, image_id, str(record_index)] + record[:2] + [str(field_index)] + values)
    _write(path, tables)


def delete_template(path, name):
    wb = _open(path)
    try:
        tables = _tables(wb)
    finally:
        wb.close()
    if not any(r[0] == name for r in tables['テンプレート設定']):
        raise TemplateError('テンプレートがありません: %s' % name)
    _write(path, {title: [r for r in rows if r[0] != name] for title, rows in tables.items()})


def _flag(text):
    if text not in ('0', '1'):
        raise TemplateError('出力・真偽値は 0 または 1 を入力してください: %r' % text)
    return text


def _number(text, minimum=1):
    try:
        value = int(text)
    except ValueError:
        raise TemplateError('整数を入力してください: %r' % text)
    if value < minimum:
        raise TemplateError('%d以上を入力してください: %r' % (minimum, text))
    return value


def load_template(path, name):
    path = Path(path)
    wb = _open(path)
    try:
        tables = {title: [r for r in rows if r[0] == name] for title, rows in _tables(wb).items()}
    finally:
        wb.close()
    settings = {}
    for _, key, _label, kind, value in tables['テンプレート設定']:
        if key in settings or kind not in ('str', 'bool'):
            raise TemplateError('設定キーの重複または型が不正です: %s' % key)
        settings[key] = (_flag(value) == '1') if kind == 'bool' else value
    if settings.pop('template_version', None) != '1':
        raise TemplateError('テンプレートがないか、バージョンが未対応です。')
    fields, seen = [], set()
    for row in tables['編集項目']:
        data = dict(zip(FIELD_KEYS, row[1:]))
        include = _flag(data.pop('include'))
        for key in ('row_number', 'occurrence_index', 'occurrence_count'):
            data[key] = _number(data[key])
        data['max_digits'] = None if data['max_digits'] == '' else _number(data['max_digits'], 0)
        field = LayoutField(**data)
        if field.instance_key in seen:
            raise TemplateError('編集項目の元Excel行・展開番号が重複しています。')
        seen.add(field.instance_key)
        if field.value == '':
            field.attribute_flag = field.coordinates = ''
        fields.append((field, include))
    records = {}
    for row in tables['出力OCR']:
        _, image_id, order, form, presence, position, fid, value, attr, coord = row
        key = (image_id, _number(order))
        position = _number(position, 0)
        if position == 0 and any((fid, value, attr, coord)):
            raise TemplateError('項目順0は項目なしの帳票専用です。値を設定する場合は1から採番してください。')
        group = records.setdefault(key, (form, presence, {}))
        if group[:2] != (form, presence) or position in group[2]:
            raise TemplateError('出力OCRの帳票情報・項目順が重複または不一致です。')
        group[2][position] = [fid, value, attr if value else '', coord if value else '']
    packages, image_ids = [], set()
    for row in tables['出力帳票']:
        _, image_id, include, front, back, has_txt, extra = row[:7]
        if not image_id or image_id in image_ids:
            raise TemplateError('画像IDが空または重複しています。')
        image_ids.add(image_id)
        include, has_txt = _flag(include), _flag(has_txt)
        kwargs = dict(zip(PACKAGE_KEYS, row[7:]))
        for side, relative in (('front', front), ('back', back)):
            if not relative:
                if side == 'front':
                    raise TemplateError('正面画像を指定してください。')
                continue
            asset = Path(relative)
            if asset.is_absolute() or '..' in asset.parts:
                raise TemplateError('画像はテンプレートからの相対パスを指定してください。')
            asset = path.parent / asset
            if not asset.is_file():
                raise TemplateError('画像資材がありません: %s' % asset)
            kwargs[side + '_image_bytes'] = asset.read_bytes()
            kwargs[side + '_extension'] = asset.suffix
        groups = sorted((k[1], v) for k, v in records.items() if k[0] == image_id)
        if has_txt == '0' and groups:
            raise TemplateError('TXTなしの画像に出力OCRが設定されています。')
        text = None
        if has_txt == '1':
            if not groups or [g[0] for g in groups] != list(range(1, len(groups) + 1)):
                raise TemplateError('出力OCRの帳票順は1から連続させてください。')
            stream = io.StringIO()
            writer = csv.writer(stream, quoting=csv.QUOTE_ALL, lineterminator='\r\n')
            for _, (form, presence, values) in groups:
                positions = sorted(values)
                if positions != [0] and positions != list(range(1, len(values) + 1)):
                    raise TemplateError('出力OCRの項目順は1から連続させてください。')
                record = [form, presence]
                for position in positions:
                    if position:
                        record.extend(values[position])
                writer.writerow(record)
            text = stream.getvalue()
            _form_records(text)
        kwargs['front_recognition_text'] = text
        kwargs['extra_fields'] = parse_extra_fields(extra)
        kwargs['source_label'] = 'テンプレート: ' + name
        packages.append((PackageItem(**kwargs), include))
    if any(key[0] not in image_ids for key in records):
        raise TemplateError('出力OCRの画像IDが出力帳票にありません。')
    return settings, fields, packages
