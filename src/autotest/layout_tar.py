# -*- coding: utf-8 -*-
"""正面・背面画像、認識結果TXT、任意CSVを1つのTARへまとめる。

正面の画像/TXTは末尾F、任意の背面画像/TXTは末尾Rを共有する。正面TXTは
既存のFORM生成内容を保持し、背面TXTだけを編集可能な1フィールドとして扱う。
"""

import csv
import io
import os
import re
import shutil
import tarfile
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple


class LayoutTarError(Exception):
    """出力リストまたはTAR設定に問題がある。"""


INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
IMAGE_EXTENSIONS = (".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp", ".gif")


def _safe_name(value: str, fallback: str) -> str:
    name = INVALID_FILENAME_CHARS.sub("_", str(value or "")).strip(" .")
    if not name:
        name = fallback
    upper = name.upper()
    if upper in ("CON", "PRN", "AUX", "NUL") or re.match(r"^(COM|LPT)[1-9]$", upper):
        name = "_" + name
    return name[:180]


def _normalise_extension(value: str) -> str:
    raw = str(value or ".tif").strip()
    if not raw.startswith("."):
        raw = "." + raw
    if raw.lower() not in IMAGE_EXTENSIONS:
        raise LayoutTarError(
            "画像拡張子は %s のいずれかにしてください: %s"
            % (" / ".join(IMAGE_EXTENSIONS), value))
    return raw


def base_name_from_front(path: Path) -> str:
    """`ABC_F.tif` / `ABCF.tif` の正面末尾Fを除いた基礎名を返す。"""
    stem = Path(path).stem
    return stem[:-1] if stem.upper().endswith("F") and len(stem) > 1 else stem


def _case_insensitive_sibling(path: Path, wanted_name: str) -> Optional[Path]:
    parent = Path(path).parent
    direct = parent / wanted_name
    if direct.is_file():
        return direct
    if not parent.is_dir():
        return None
    folded = wanted_name.lower()
    try:
        return next((candidate for candidate in parent.iterdir()
                     if candidate.is_file() and candidate.name.lower() == folded), None)
    except OSError:
        return None


def matching_back_image(front_path: Path) -> Optional[Path]:
    """正面画像と同じ基礎名・拡張子で末尾Rの背面画像を探す。"""
    path = Path(front_path)
    base = base_name_from_front(path)
    return _case_insensitive_sibling(path, base + "R" + path.suffix)


def matching_recognition_file(image_path: Path) -> Optional[Path]:
    """画像と同じ側別名のTXTを探し、正面だけ旧形式名にも対応する。"""
    path = Path(image_path)
    exact = _case_insensitive_sibling(path, path.stem + ".txt")
    if exact is not None:
        return exact
    if path.stem.upper().endswith("F"):
        return _case_insensitive_sibling(path, base_name_from_front(path) + ".txt")
    return None


class PackageItem(object):
    """TARへ入れる1件分の画像（正面必須、背面任意）と対応情報。"""

    def __init__(self, base_name: str, form_id: str = "",
                 front_recognition_text: Optional[str] = None,
                 back_recognition_result: str = "1",
                 front_image_path: Optional[Path] = None,
                 front_image_bytes: Optional[bytes] = None,
                 front_extension: str = ".tif",
                 back_image_path: Optional[Path] = None,
                 back_image_bytes: Optional[bytes] = None,
                 back_extension: str = "",
                 related_file: str = "",
                 extra_fields: Optional[Dict[str, str]] = None,
                 source_label: str = "",
                 scan_batch_id: str = "",
                 image_sequence: str = "",
                 arrival_date: str = "",
                 application_number: str = "",
                 reception_number: str = "",
                 format_id: str = "",
                 delivery_date: str = "",
                 delivery_shot: str = "") -> None:
        self.base_name = str(base_name or "").strip()
        self.form_id = str(form_id or "").strip()
        self.front_recognition_text = (
            None if front_recognition_text is None else str(front_recognition_text))
        self.back_recognition_result = str(
            back_recognition_result if back_recognition_result is not None else "")
        self.front_image_path = Path(front_image_path) if front_image_path is not None else None
        self.front_image_bytes = front_image_bytes
        self.front_extension = _normalise_extension(front_extension)
        self.back_image_path = Path(back_image_path) if back_image_path is not None else None
        self.back_image_bytes = back_image_bytes
        self.back_extension = _normalise_extension(back_extension or self.front_extension)
        self.related_file = str(related_file or "").strip()
        self.extra_fields = {
            str(key).strip(): str(value if value is not None else "")
            for key, value in (extra_fields or {}).items() if str(key).strip()
        }
        self.source_label = str(source_label or "").strip()
        self.scan_batch_id = str(scan_batch_id or "").strip()
        self.image_sequence = str(image_sequence or "").strip()
        self.arrival_date = str(arrival_date or "").strip()
        self.application_number = str(application_number or "").strip()
        self.reception_number = str(reception_number or "").strip()
        self.format_id = str(format_id or "").strip()
        self.delivery_date = str(delivery_date or "").strip()
        self.delivery_shot = str(delivery_shot or "").strip()

    @property
    def safe_base_name(self) -> str:
        return _safe_name(self.base_name, "image")

    @property
    def front_image_name(self) -> str:
        return self.safe_base_name + "F" + self.front_extension

    @property
    def has_back_image(self) -> bool:
        return self.back_image_bytes is not None or self.back_image_path is not None

    @property
    def back_image_name(self) -> str:
        return self.safe_base_name + "R" + self.back_extension

    @property
    def front_recognition_name(self) -> str:
        return self.safe_base_name + "F.txt"

    @property
    def back_recognition_name(self) -> str:
        return self.safe_base_name + "R.txt"

    @property
    def has_front_recognition(self) -> bool:
        return self.front_recognition_text is not None

    @staticmethod
    def _payload(path: Optional[Path], payload: Optional[bytes], label: str) -> bytes:
        if payload is not None:
            return bytes(payload)
        if path is None:
            raise LayoutTarError("%s画像が指定されていません" % label)
        if not path.is_file():
            raise LayoutTarError("%s画像ファイルが見つかりません: %s" % (label, path))
        try:
            return path.read_bytes()
        except OSError as exc:
            raise LayoutTarError("%s画像ファイルを読めません: %s: %s" % (label, path, exc))

    def front_payload(self) -> bytes:
        return self._payload(self.front_image_path, self.front_image_bytes, "正面")

    def back_payload(self) -> bytes:
        return self._payload(self.back_image_path, self.back_image_bytes, "背面")

    def manifest_values(self, include_recognition_txt: bool) -> Dict[str, str]:
        values = {
            "image_file": self.front_image_name,
            "front_image_file": self.front_image_name,
            "back_image_file": self.back_image_name if self.has_back_image else "",
            "recognition_file": (
                self.front_recognition_name
                if include_recognition_txt and self.has_front_recognition else ""),
            "front_recognition_file": (
                self.front_recognition_name
                if include_recognition_txt and self.has_front_recognition else ""),
            "back_recognition_file": (
                self.back_recognition_name
                if include_recognition_txt and self.has_back_image else ""),
            "related_file": self.related_file or self.front_image_name,
            "form_id": self.form_id,
            "recognition_result": (
                self.back_recognition_result if self.has_back_image else ""),
            "back_recognition_result": (
                self.back_recognition_result if self.has_back_image else ""),
            "base_name": self.safe_base_name,
            "source_file": (
                self.front_image_path.name if self.front_image_path is not None
                else (self.source_label or "generated")),
            "back_source_file": (
                self.back_image_path.name if self.back_image_path is not None
                else (self.source_label if self.has_back_image else "")),
            "scan_batch_id": self.scan_batch_id,
            "image_sequence": self.image_sequence,
            "arrival_date": self.arrival_date,
            "application_number": self.application_number,
            "reception_number": self.reception_number,
            "format_id": self.format_id,
            "delivery_date": self.delivery_date,
            "delivery_shot": self.delivery_shot,
        }
        values.update(self.extra_fields)
        return values

    def image_list_values(self, image_name: str, sequence: str) -> List[str]:
        return [
            self.scan_batch_id,
            self.image_sequence or sequence,
            image_name,
            self.arrival_date,
            self.form_id,
            self.application_number,
            self.reception_number,
            self.format_id,
            self.delivery_date,
            self.delivery_shot,
        ]


class PackageResult(object):
    def __init__(self, tar_file: Path, archive_members: List[str],
                 item_count: int, manifest_name: str = "",
                 view_folder: Optional[Path] = None) -> None:
        self.tar_file = tar_file
        self.archive_members = archive_members
        self.item_count = item_count
        self.manifest_name = manifest_name
        self.view_folder = view_folder


def format_package_tar_name(image_txt_template: str,
                            image_csv_template: str,
                            include_manifest: bool,
                            source: str,
                            form_ids: Sequence[str]) -> str:
    raw = image_csv_template if include_manifest else image_txt_template
    raw = str(raw or "").strip()
    if not raw:
        raw = "image_package" if include_manifest else "{source}_layout_data"
    values = []
    for value in form_ids:
        value = str(value or "").strip()
        if value and value not in values:
            values.append(value)
    form_value = values[0] if len(values) == 1 else "all"
    try:
        return raw.format(source=str(source or "images"), form_id=form_value)
    except (KeyError, ValueError, IndexError) as exc:
        raise LayoutTarError("梱包TAR名テンプレートが不正です: %s" % exc)


def parse_extra_fields(text: str) -> Dict[str, str]:
    """`key=value; key2=value2` をCSV追加フィールドへ変換する。"""
    result = {}  # type: Dict[str, str]
    raw = str(text or "").strip()
    if not raw:
        return result
    for part in raw.split(";"):
        part = part.strip()
        if not part:
            continue
        if "=" not in part:
            raise LayoutTarError(
                "CSV追加項目は key=value を ; で区切ってください: %s" % part)
        key, value = part.split("=", 1)
        key = key.strip()
        if not key:
            raise LayoutTarError("CSV追加項目の項目名が空です: %s" % part)
        if key in result:
            raise LayoutTarError("CSV追加項目が重複しています: %s" % key)
        result[key] = value.strip()
    return result


def format_extra_fields(values: Dict[str, str]) -> str:
    return "; ".join("%s=%s" % (key, value) for key, value in values.items())


def parse_manifest_columns(text: str) -> List[Tuple[str, str]]:
    """CSV列指定を `(見出し, 値キー)` の配列へ変換する。

    `正面画像=front_image_file,ID=form_id,case_type` のように指定できる。
    `=` が無い列は見出しと値キーを同じ文字列として扱う。
    """
    columns = []  # type: List[Tuple[str, str]]
    seen_headers = set()
    for part in str(text or "").split(","):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            header, key = (value.strip() for value in part.split("=", 1))
        else:
            header, key = part, part
        if not header or not key:
            raise LayoutTarError("CSV列は `見出し=値キー` で指定してください: %s" % part)
        folded = header.lower()
        if folded in seen_headers:
            raise LayoutTarError("CSV見出しが重複しています: %s" % header)
        seen_headers.add(folded)
        columns.append((header, key))
    if not columns:
        raise LayoutTarError("CSVを生成する場合はCSV列を1つ以上指定してください")
    return columns


def _normalise_columns(columns: Sequence) -> List[Tuple[str, str]]:
    if isinstance(columns, str):
        return parse_manifest_columns(columns)
    result = []  # type: List[Tuple[str, str]]
    for spec in columns:
        if isinstance(spec, str):
            result.extend(parse_manifest_columns(spec))
        elif isinstance(spec, (tuple, list)) and len(spec) == 2:
            result.append((str(spec[0]).strip(), str(spec[1]).strip()))
        else:
            raise LayoutTarError("CSV列指定が不正です: %r" % (spec,))
    if not result:
        raise LayoutTarError("CSVを生成する場合はCSV列を1つ以上指定してください")
    seen = set()
    for header, key in result:
        if not header or not key:
            raise LayoutTarError("CSV列の見出しと値キーは空にできません")
        folded = header.lower()
        if folded in seen:
            raise LayoutTarError("CSV見出しが重複しています: %s" % header)
        seen.add(folded)
    return result


def _manifest_payload(items: Sequence[PackageItem], columns: Sequence,
                      encoding: str, include_recognition_txt: bool) -> bytes:
    normalised = _normalise_columns(columns)
    stream = io.StringIO()
    writer = csv.writer(stream, lineterminator="\r\n")
    writer.writerow([header for header, _key in normalised])
    for item in items:
        values = item.manifest_values(include_recognition_txt)
        writer.writerow([values.get(key, "") for _header, key in normalised])
    try:
        return stream.getvalue().encode(encoding)
    except (LookupError, UnicodeEncodeError) as exc:
        raise LayoutTarError("CSVを%sでエンコードできません: %s" % (encoding, exc))


def _image_list_payload(items: Sequence[PackageItem], encoding: str) -> bytes:
    stream = io.StringIO()
    writer = csv.writer(stream, lineterminator="\r\n", quoting=csv.QUOTE_ALL)
    for index, item in enumerate(items, 1):
        sequence = "%03d" % index
        rows = [item.image_list_values(item.front_image_name, sequence)]
        if item.has_back_image:
            rows.append(item.image_list_values(item.back_image_name, sequence))
        for row in rows:
            if any("\r" in value or "\n" in value for value in row):
                raise LayoutTarError("固定10列CSVの値には改行を使用できません")
            writer.writerow(row)
    try:
        return stream.getvalue().encode(encoding)
    except (LookupError, UnicodeEncodeError) as exc:
        raise LayoutTarError("CSVを%sでエンコードできません: %s" % (encoding, exc))


def _add_tar_member(archive, name: str, payload: bytes) -> None:
    info = tarfile.TarInfo(name=name)
    info.size = len(payload)
    info.mtime = 0
    info.mode = 0o644
    info.uid = 0
    info.gid = 0
    info.uname = ""
    info.gname = ""
    archive.addfile(info, io.BytesIO(payload))


def _encode_text_file(name: str, value: str, encoding: str) -> bytes:
    text = str(value).replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
    try:
        return (text.replace("\n", "\r\n") + "\r\n").encode(encoding)
    except (LookupError, UnicodeEncodeError) as exc:
        raise LayoutTarError(
            "%s を%sでエンコードできません: %s" % (name, encoding, exc))


def _encode_single_field(name: str, value: str, encoding: str) -> bytes:
    if "\r" in str(value) or "\n" in str(value):
        raise LayoutTarError("背面認識値は1フィールドのため改行できません: %s" % name)
    stream = io.StringIO()
    csv.writer(
        stream, lineterminator="\r\n", quoting=csv.QUOTE_ALL).writerow([str(value)])
    try:
        return stream.getvalue().encode(encoding)
    except (LookupError, UnicodeEncodeError) as exc:
        raise LayoutTarError(
            "%s を%sでエンコードできません: %s" % (name, encoding, exc))


def _atomic_write(path: Path, payload: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(
        prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(handle, "wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_name, str(path))
    except Exception:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise


def _path_exists(path: Path) -> bool:
    return os.path.lexists(str(path))


def _remove_path(path: Path) -> None:
    if not _path_exists(path):
        return
    if path.is_symlink() or path.is_file():
        path.unlink()
    else:
        shutil.rmtree(str(path))


def _publish_package_outputs(output_dir: Path, tar_file: Path,
                             tar_payload: bytes,
                             rendered: Sequence[Tuple[str, bytes]],
                             view_folder: Optional[Path],
                             overwrite: bool) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    targets = [tar_file]
    if view_folder is not None:
        targets.append(view_folder)

    existing = [path for path in targets if _path_exists(path)]
    if existing and not overwrite:
        raise LayoutTarError(
            "既存のTARまたは確認用フォルダを上書きしません。"
            "上書きを有効にするかTAR名を変えてください: %s"
            % ", ".join(str(path) for path in existing))
    if _path_exists(tar_file) and tar_file.is_dir():
        raise LayoutTarError("TAR出力先がフォルダです: %s" % tar_file)
    if (view_folder is not None and _path_exists(view_folder) and
            (view_folder.is_symlink() or not view_folder.is_dir())):
        raise LayoutTarError("確認用フォルダの出力先がフォルダではありません: %s" % view_folder)

    staging_root = Path(tempfile.mkdtemp(
        prefix=".%s." % tar_file.stem, dir=str(output_dir)))
    backup_root = None  # type: Optional[Path]
    backups = {}  # type: Dict[Path, Path]
    published = []  # type: List[Path]
    try:
        staged_tar = staging_root / tar_file.name
        _atomic_write(staged_tar, tar_payload)
        staged_outputs = [(staged_tar, tar_file)]
        if view_folder is not None:
            staged_folder = staging_root / view_folder.name
            staged_folder.mkdir()
            for name, payload in rendered:
                _atomic_write(staged_folder / name, payload)
            staged_outputs.append((staged_folder, view_folder))

        if existing:
            backup_root = Path(tempfile.mkdtemp(
                prefix=".%s.backup." % tar_file.stem, dir=str(output_dir)))
            for target in existing:
                backup = backup_root / target.name
                os.replace(str(target), str(backup))
                backups[target] = backup

        for staged, target in staged_outputs:
            os.replace(str(staged), str(target))
            published.append(target)
    except Exception as exc:
        for target in reversed(published):
            try:
                _remove_path(target)
            except OSError:
                pass
        restore_errors = []
        for target, backup in backups.items():
            if _path_exists(backup):
                try:
                    os.replace(str(backup), str(target))
                except OSError as restore_exc:
                    restore_errors.append(str(restore_exc))
        if restore_errors:
            kept_backup = backup_root
            backup_root = None
            raise LayoutTarError(
                "出力の復元に失敗しました。退避データを確認してください: %s: %s"
                % (kept_backup, "; ".join(restore_errors))) from exc
        raise
    finally:
        shutil.rmtree(str(staging_root), ignore_errors=True)
        if backup_root is not None:
            shutil.rmtree(str(backup_root), ignore_errors=True)


def build_image_tar(items: Sequence[PackageItem], output_dir: Path,
                    tar_name: str,
                    include_recognition_txt: bool = True,
                    include_manifest_csv: bool = False,
                    manifest_name: str = "file_list.csv",
                    manifest_columns: Optional[Sequence] = None,
                    text_encoding: str = "cp932",
                    csv_encoding: str = "cp932",
                    overwrite: bool = False,
                    manifest_style: str = "custom",
                    create_view_folder: bool = True) -> PackageResult:
    """正面、任意の背面、任意の認識TXT/一覧CSVを1つのTARへまとめる。"""
    selected = list(items)
    if not selected:
        raise LayoutTarError("TARへ入れる画像がありません")

    rendered = []  # type: List[Tuple[str, bytes]]
    seen = set()
    for item in selected:
        members = [(item.front_image_name, item.front_payload())]
        if include_recognition_txt:
            if not item.has_front_recognition:
                raise LayoutTarError(
                    "正面TXTがありません。FORMから追加するか対応TXTを配置してください: %s"
                    % item.front_image_name)
            members.append((
                item.front_recognition_name,
                _encode_text_file(
                    item.front_recognition_name,
                    item.front_recognition_text or "", text_encoding)))
        if item.has_back_image:
            members.append((item.back_image_name, item.back_payload()))
            if include_recognition_txt:
                members.append((
                    item.back_recognition_name,
                    _encode_single_field(
                        item.back_recognition_name,
                        item.back_recognition_result, text_encoding)))
        for name, payload in members:
            folded = name.lower()
            if folded in seen:
                raise LayoutTarError("TAR内のファイル名が重複します: %s" % name)
            seen.add(folded)
            rendered.append((name, payload))

    actual_manifest_name = ""
    if include_manifest_csv:
        stem = str(manifest_name or "file_list.csv").strip()
        if stem.lower().endswith(".csv"):
            stem = stem[:-4]
        actual_manifest_name = _safe_name(stem, "file_list") + ".csv"
        folded = actual_manifest_name.lower()
        if folded in seen:
            raise LayoutTarError(
                "CSV名が画像/TXT名と重複します: %s" % actual_manifest_name)
        style = str(manifest_style or "custom").strip().lower()
        if style == "image_list":
            manifest_payload = _image_list_payload(selected, csv_encoding)
        elif style == "custom":
            columns = manifest_columns or (
                ("front_image_file", "front_image_file"),
                ("back_image_file", "back_image_file"),
                ("form_id", "form_id"),
                ("back_recognition_result", "back_recognition_result"),
                ("related_file", "related_file"),
            )
            manifest_payload = _manifest_payload(
                selected, columns, csv_encoding, include_recognition_txt)
        else:
            raise LayoutTarError("CSV形式は image_list / custom から選択してください: %s" % style)
        rendered.append((
            actual_manifest_name,
            manifest_payload))
        seen.add(folded)

    raw_tar_name = str(tar_name or "image_package").strip()
    if raw_tar_name.lower().endswith(".tar"):
        raw_tar_name = raw_tar_name[:-4]
    output_dir = Path(output_dir)
    safe_tar_stem = _safe_name(raw_tar_name, "image_package")
    tar_file = output_dir / (safe_tar_stem + ".tar")
    view_folder = output_dir / safe_tar_stem if create_view_folder else None

    stream = io.BytesIO()
    with tarfile.open(fileobj=stream, mode="w") as archive:
        for name, payload in rendered:
            _add_tar_member(archive, name, payload)
    _publish_package_outputs(
        output_dir, tar_file, stream.getvalue(), rendered,
        view_folder, overwrite)
    return PackageResult(
        tar_file=tar_file,
        archive_members=[name for name, _payload in rendered],
        item_count=len(selected), manifest_name=actual_manifest_name,
        view_folder=view_folder)
