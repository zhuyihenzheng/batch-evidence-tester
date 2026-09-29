# -*- coding: utf-8 -*-
"""受領 TXT を定義と照合し、原値を保持した確認用 Excel を作る。"""

import argparse
import csv
import hashlib
import io
import os
import re
import sys
import tempfile
import zipfile
from collections import Counter, OrderedDict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .layout_txt import LayoutTxtError, read_layout_fields


SCOPE = ("Excel定義に対応する実値を表示し、項目数・順序・文字数などの注意点を参考情報として記載します。"
         "座標は任意です。空欄でも注意点にしません。業務内容は実値と定義を見て確認してください。")
DETAIL_HEADERS = [
    "ファイル", "レコード", "物理開始行", "FORM_ID", "対象有無", "TXT項目順",
    "FieldID", "項目名", "受領OCR値", "文字数", "最大桁数", "属性", "座標",
    "参考情報", "注意点", "定義シート", "定義行", "LAYOUT_ID", "ELEMENT_ID",
    "展開番号", "データ型", "IME", "入力属性", "入力規則", "補足", "出力例",
    "確認内容（手入力）", "確認者", "確認メモ",
]


def _sha(data):
    return hashlib.sha256(data).hexdigest()


class Inspection(object):
    def __init__(self, definition, sheet, header, columns, definition_hash, encoding):
        self.definition = str(definition)
        self.sheet = sheet
        self.header = header
        self.columns = columns
        self.definition_hash = definition_hash
        self.encoding = encoding
        self.created = datetime.now().isoformat(timespec="seconds")
        self.files = []
        self.records = []
        self.details = []

    @property
    def exit_code(self):
        # 参考情報の有無は、読込・Excel出力の成功とは別に扱う。
        return 0


def _detail(result, record, position, block, field, issues, status):
    field_id, value, attribute, coordinates = (block + [None] * 4)[:4]
    return [
        record["file"], record["number"], record["line"], record["form"],
        record["presence"], position, field_id,
        field.item_name if field else "", value, len(value) if value is not None else None,
        field.max_digits if field else None, attribute, coordinates, status,
        " / ".join(issues), result.sheet if field else "", field.row_number if field else None,
        field.layout_id if field else "", field.source_element_id if field else "",
        ("%d/%d" % (field.occurrence_index, field.occurrence_count)) if field else "",
        field.data_type if field else "", field.ime_name if field else "",
        field.input_attribute if field else "", field.input_rule if field else "",
        field.notes if field else "", field.output_example if field else "", "", "", "",
    ]


def _inspect_record(result, groups, record, values, parse_error, block_width=4):
    issues = []
    review = []
    record.update(form=values[0] if values else "",
                  presence=values[1] if len(values) > 1 else "")
    fields = groups.get(record["form"])
    if parse_error:
        issues.append(parse_error)
    if len(values) < 2:
        issues.append("FORM_ID / 対象有無が不足")
    if fields is None:
        issues.append("FORM_IDが定義外または空")
    if record["presence"] not in ("0", "1"):
        issues.append("対象有無は0または1が必要")
    elif record["presence"] == "0":
        review.append("対象有無=0。項目の要否を確認")
    blocks = [values[index:index + block_width] for index in range(2, len(values), block_width)]
    counts = Counter(block[0] for block in blocks)
    lookup = {field.field_id: field for field in fields or []}
    if fields and len(blocks) != len(fields):
        issues.append("項目数不一致: 定義%d / 受領%d" % (len(fields), len(blocks)))
    if fields and [block[0] for block in blocks] != [field.field_id for field in fields]:
        issues.append("FieldIDの並びが定義と不一致")
    detail_statuses = []
    for index, block in enumerate(blocks, 1):
        field = lookup.get(block[0])
        problems = []
        cautions = []
        if field is None:
            problems.append("FieldIDが定義外または空")
        if counts[block[0]] > 1:
            problems.append("FieldIDが重複")
        if len(block) < 3:
            problems.append("項目ブロックのOCR値または属性が不足")
        value = block[1] if len(block) > 1 else None
        if value is not None and field and field.max_digits is not None:
            if len(value) > field.max_digits:
                problems.append("最大桁数超過（文字数基準）")
        if value == "":
            cautions.append("OCR値が空。原票を確認")
        if len(block) > 2:
            if block[2] not in ("0", "1", "2", "1,2"):
                problems.append("属性が定義外")
            elif block[2] != "0":
                cautions.append("属性%s（1:個数不正 / 2:認識不可）" % block[2])
        if len(block) > 3 and block[3] and not re.fullmatch(r"[0-9]+,[0-9]+,[0-9]+,[0-9]+", block[3]):
            cautions.append("座標の形式を確認（参考:4個の非負整数）")
        if fields and index <= len(fields) and block[0] != fields[index - 1].field_id:
            problems.append("FieldIDの順序が定義と不一致")
        status = "注意点あり" if problems or cautions else ""
        detail_statuses.append(status)
        result.details.append(_detail(result, record, index, block, field,
                                      problems + cautions, status))
    for field in fields or []:
        if field.field_id not in counts:
            result.details.append(_detail(result, record, None, [field.field_id], field,
                                          ["TXTに項目なし（定義から補記）"], "注意点あり"))
            detail_statuses.append("注意点あり")
    if "注意点あり" in detail_statuses:
        review.append("項目明細に注意点あり")
    record.update(status="注意点あり" if issues or review else "",
                  issues=" / ".join(issues + review), expected=len(fields) if fields else None,
                  actual=len(blocks))
    result.records.append(record)


def inspect_txt(excel_path, txt_paths, encoding="cp932", block_width=4, **definition_options):
    """同一FORMの複数レコードも別々に照合し、実値を補正しない。"""
    definition = Path(excel_path).resolve()
    paths = [Path(path).resolve() for path in txt_paths]
    if not paths:
        raise LayoutTxtError("受領TXTを1件以上選択してください。")
    if len(set(paths)) != len(paths):
        raise LayoutTxtError("同じ受領TXTが重複選択されています。")
    if block_width not in (3, 4):
        raise LayoutTxtError("項目形式は座標列あり（4要素）または座標列なし（3要素）にしてください。")
    # 生成用既定値・画面編集は実データの解釈に混ぜない。
    definition_options.update(default_value_column="none", coordinates_column="none",
                              profile="normal", date_mode="wareki", coverage_form_id="")
    try:
        before = definition.read_bytes()
        fields, sheet, header, columns = read_layout_fields(definition, **definition_options)
        if definition.read_bytes() != before:
            raise LayoutTxtError("読込中に定義Excelが変更されました。再実行してください。")
    except OSError as exc:
        raise LayoutTxtError("定義Excelを読めません: %s" % exc)
    result = Inspection(definition, sheet, header, columns, _sha(before), encoding)
    result.block_width = block_width
    groups = OrderedDict()
    for field in fields:
        groups.setdefault(field.form_id, []).append(field)
    for path in paths:
        try:
            payload = path.read_bytes()
            text = payload.decode("utf-8" if encoding == "utf-8-sig" else encoding)
        except (OSError, UnicodeError, LookupError) as exc:
            raise LayoutTxtError("受領TXTを読めません（文字コード %s）: %s: %s" %
                                 (encoding, path, exc))
        result.files.append((str(path), len(payload), _sha(payload)))
        # BOMは解析時だけ除去する。証跡の原文とハッシュには保持する。
        raw_lines = io.StringIO(text, newline="").readlines()
        parsed_text = text[1:] if text.startswith("\ufeff") else text
        reader = csv.reader(io.StringIO(parsed_text, newline=""), strict=True)
        previous = 0
        number = 0
        while True:
            error = ""
            try:
                values = next(reader)
            except StopIteration:
                break
            except csv.Error as exc:
                # 不正引用符で境界を失った後は推測して次レコードに割り当てない。
                values = []
                error = "CSV解析不可: %s（以降の原文を保持）" % exc
            number += 1
            end = len(raw_lines) if error else reader.line_num
            record = dict(file=str(path), number=number, line=previous + 1,
                          end_line=end, raw="".join(raw_lines[previous:end]))
            _inspect_record(result, groups, record, values, error, block_width=block_width)
            previous = end
            if error:
                break
        if not number:
            _inspect_record(result, groups, dict(file=str(path), number=1, line=1,
                                                end_line=1, raw=text), [], "空ファイル", block_width=block_width)
    return result


def _excel_text(value):
    # XMLに格納できない制御文字は可視化する。通常の値は変更しない。
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]",
                  lambda match: "\\u%04X" % ord(match.group()), value)


def _append(ws, values):
    if ws.max_row >= 1048576:
        raise LayoutTxtError("Excelの行数上限を超えました。TXTを分割して実行してください。")
    row = ws.max_row + 1 if ws.max_row > 1 or ws.cell(1, 1).value is not None else 1
    for col, value in enumerate(values, 1):
        if isinstance(value, str):
            value = _excel_text(value)
            if len(value) > 32767:
                raise LayoutTxtError("Excelのセル文字数上限を超えました: %s 行%d 列%d" %
                                     (ws.title, row, col))
        cell = ws.cell(row, col, value)
        if isinstance(value, str):
            # =始まりや先頭ゼロも受領文字列そのもの。数式として実行しない。
            cell.data_type = "s"
            cell.number_format = "@"


def _style(ws, widths, status_column=None):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="244660")
        cell.font = Font(name="Meiryo", size=10, color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Meiryo", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[row[0].row].height = 45
        if status_column:
            cell = row[status_column - 1]
            if cell.value == "注意点あり":
                cell.fill = PatternFill("solid", fgColor="FFF2CC")


def export_inspection(result, output_path, overwrite=False):
    """原本を保護し、保存失敗時にも既存成果物を壊さない。"""
    path = Path(output_path).resolve()
    sources = [Path(result.definition)] + [Path(item[0]) for item in result.files]
    if path in sources or (path.exists() and any(os.path.samefile(str(path), str(p))
                                               for p in sources if p.exists())):
        raise LayoutTxtError("出力先に入力原本は指定できません。")
    if path.suffix.lower() != ".xlsx":
        raise LayoutTxtError("出力先は .xlsx にしてください。")
    if path.exists() and not overwrite:
        raise LayoutTxtError("同名成果物があります。別名にするか上書きを指定してください。")
    wb = Workbook()
    summary = wb.active
    summary.title = "受領確認"
    _append(summary, ["受領TXT確認", "内容"])
    for pair in [("解析日時", result.created), ("定義Excel", result.definition),
                 ("定義SHA-256", result.definition_hash), ("定義シート", result.sheet),
                 ("見出し行", result.header), ("文字コード", result.encoding),
                 ("ファイル数", len(result.files)), ("レコード数", len(result.records)),
                 ("注意点のあるレコード数", sum(bool(r["issues"]) for r in result.records)),
                 ("TXT項目形式", "座標列あり（空欄可）" if result.block_width == 4 else "座標列なし"),
                 ("参考情報の範囲", SCOPE),
                 ("FieldID対応", "各FORM内の展開後連番。カレンダーは46項目。ELEMENT_IDとは別。"),
                 ("原値の表示", "空白・先頭ゼロを保持。XML禁止制御文字のみ\\uXXXXで表示。"
                  "原文は10000文字単位で分割。長文は数式バーで全文確認。"),
                 ("列設定", ", ".join("%s=%s" % (k, v) for k, v in sorted(result.columns.items())))]:
        _append(summary, pair)
    for filename, size, digest in result.files:
        _append(summary, ["受領ファイル", filename])
        _append(summary, ["バイト数", size])
        _append(summary, ["SHA-256", digest])
    _style(summary, [26, 115])
    summary.auto_filter.ref = None
    summary.row_dimensions[12].height = 64
    records = wb.create_sheet("レコード一覧")
    _append(records, ["ファイル", "レコード", "物理開始行", "物理終了行", "FORM_ID", "対象有無",
                      "定義項目数", "受領項目数", "参考情報", "注意点"])
    for record in result.records:
        _append(records, [record[key] for key in (
            "file", "number", "line", "end_line", "form", "presence", "expected",
            "actual", "status", "issues")])
    _style(records, [42, 12, 12, 12, 14, 12, 14, 14, 14, 85], 9)
    details = wb.create_sheet("項目明細")
    _append(details, DETAIL_HEADERS)
    for row in result.details:
        _append(details, row)
    _style(details, [42, 12, 12, 14, 12, 12, 12, 26, 42, 10, 12, 12, 23, 14, 65,
                     20, 10, 14, 16, 12, 20, 18, 18, 32, 32, 24, 22, 18, 36], 14)
    details.freeze_panes = "I2"
    for row in details.iter_rows(min_row=2, min_col=27, max_col=29):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor="FFF2CC")
    raw = wb.create_sheet("受領原文")
    _append(raw, ["ファイル", "レコード", "物理開始行", "分割番号", "原文（分割順・制御文字は可視化）"])
    for record in result.records:
        chunks = [record["raw"][i:i + 10000] for i in range(0, len(record["raw"]), 10000)] or [""]
        for index, chunk in enumerate(chunks, 1):
            # 制御文字の可視化後もセル上限を守る。
            visible = _excel_text(chunk)
            if len(visible) > 32767:
                raise LayoutTxtError("原文の制御文字が多すぎるためExcelに格納できません。")
            _append(raw, [record["file"], record["number"], record["line"], index, visible])
    _style(raw, [42, 12, 12, 12, 115])
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=".inspection_", suffix=".xlsx", dir=str(path.parent))
    os.close(fd)
    try:
        packed = io.BytesIO()
        wb.save(packed)
        packed.seek(0)
        # ElementTree系のopenpyxlはCRを実文字でXML化する場合がある。
        # XML読込時の改行正規化で受領値が変わらないよう文字参照にする。
        with zipfile.ZipFile(packed, "r") as source_zip, zipfile.ZipFile(temp_name, "w") as target_zip:
            for entry in source_zip.infolist():
                data = source_zip.read(entry.filename)
                if entry.filename.endswith(".xml"):
                    data = data.replace(b"\r", b"&#13;")
                target_zip.writestr(entry, data)
        if overwrite:
            os.replace(temp_name, str(path))
        else:
            # x モードにより保存直前の同名ファイル作成でも上書きしない。
            with open(temp_name, "rb") as src, path.open("xb") as dst:
                try:
                    import shutil
                    shutil.copyfileobj(src, dst)
                except Exception:
                    dst.close()
                    path.unlink()
                    raise
    finally:
        wb.close()
        if os.path.exists(temp_name):
            os.unlink(temp_name)
    return path


def main(argv=None):
    parser = argparse.ArgumentParser(description="受領TXTを定義Excelで照合して確認用Excelに出力")
    parser.add_argument("excel", help="レイアウト定義Excel")
    parser.add_argument("txt", nargs="+", help="受領TXT（複数指定可）")
    parser.add_argument("--out", required=True, help="確認用.xlsx出力先")
    parser.add_argument("--encoding", default="cp932", choices=("cp932", "utf-8", "utf-8-sig"))
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--header-row", type=int, default=None)
    for role, default in (("form", "auto"), ("layout", "auto"), ("field", "auto"),
                          ("item", "auto"), ("data-type", "I"), ("ime", "J"),
                          ("max-digits", "K")):
        parser.add_argument("--%s-column" % role, default=default)
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--no-coordinates", action="store_true", help="座標列なしの3要素形式で読む")
    args = parser.parse_args(argv)
    options = {role + "_column": getattr(args, role + "_column") for role in
               ("form", "layout", "field", "item", "data_type", "ime", "max_digits")}
    try:
        result = inspect_txt(args.excel, args.txt, encoding=args.encoding,
                             block_width=3 if args.no_coordinates else 4,
                             sheet_name=args.sheet, header_row=args.header_row, **options)
        path = export_inspection(result, args.out, overwrite=args.overwrite)
    except (LayoutTxtError, OSError, ValueError) as exc:
        print("[設定エラー] %s" % exc, file=sys.stderr)
        return 2
    print("確認用Excel: %s / レコード数: %d" % (path, len(result.records)))
    print(SCOPE)
    return result.exit_code


if __name__ == "__main__":
    sys.exit(main())
