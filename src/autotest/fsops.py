"""入出力フォルダの操作: テストデータ投入・処理結果の回収・フォルダ一覧取得。

安全上の約束:
  - クリア対象は settings.paths に定義された論理名のみ受け付ける。
    生パスを直接消せないようにして、設定ミスで無関係なフォルダを消す事故を防ぐ。
  - クリアはフォルダ自体を消さずに中身だけ削除する（batch 側が存在を前提にするため）。
"""

import fnmatch
import hashlib
import os
import shutil
import stat
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from .config import ConfigError, Settings


def _force_remove(func, path, exc):
    """読み取り専用属性で削除に失敗した場合だけ属性を落として再試行する。

    Windows では batch が出力したファイルに読み取り専用が付くことがあり、
    そのままでは shutil.rmtree が失敗するため、その救済に限定する。

    注意: rmtree のコールバックには os.open / os.scandir も渡ってくる。
    これらは func(path) の形では呼べないので、無条件に再試行してはいけない。
    権限以外の失敗（TypeError や存在しないパス）はそのまま送出し、
    「消せなかったのに消せたことにする」状態を作らない。
    """
    if not isinstance(exc, PermissionError):
        raise exc
    if func not in (os.unlink, os.rmdir, os.remove):
        raise exc
    os.chmod(path, stat.S_IWRITE)
    func(path)


# 中身を消してはいけない場所。設定ミスで業務データや OS を壊さないための最終防衛線。
def assert_safe_to_clear(directory: Path, project_root: Path) -> None:
    """このフォルダの中身を削除してよいか検証する。危険なら ConfigError。

    論理名しか受け付けない仕組みだけでは、その論理名が C:/ を指していた場合を
    防げない。実際に解決されたパスそのものを見て判断する。
    """
    target = Path(os.path.abspath(str(directory)))
    root = Path(os.path.abspath(str(project_root)))

    def deny(reason: str) -> None:
        raise ConfigError(
            "クリア対象として危険なパスが指定されました: %s\n  理由: %s\n"
            "  settings.yaml の paths を確認してください。" % (target, reason)
        )

    if target.parent == target:
        deny("ファイルシステムのルートです")

    home = Path(os.path.abspath(os.path.expanduser("~")))
    for protected, label in ((home, "ユーザーのホームフォルダ"), (home.parent, "ホームの親フォルダ"),
                             (root, "プロジェクトルート")):
        if target == protected:
            deny(label + "そのものです")

    # プロジェクトルートやホームの「上位」を消すと巻き添えで全部消える
    for descendant, label in ((root, "プロジェクトルート"), (home, "ホームフォルダ")):
        if descendant != target and _is_relative_to(descendant, target):
            deny("%s (%s) を含む上位フォルダです" % (label, descendant))

    # ルート直下（C:\work や /data）はフォルダ階層が浅すぎて誤削除の影響が大きい
    if len(target.parts) <= 2:
        deny("階層が浅すぎます（ルート直下のフォルダは対象にできません）")


def _is_relative_to(path: Path, other: Path) -> bool:
    """path が other の配下か。Path.is_relative_to は 3.9+ のため自前で判定する。"""
    try:
        path.relative_to(other)
        return True
    except ValueError:
        return False


def _rmtree(path: Path) -> None:
    """読み取り専用対策付きの rmtree。

    Python 3.12 で onerror が非推奨になり onexc に置き換わったため、
    利用可能な方を選ぶ（将来の 3.x で onerror が削除されても動くように）。
    """
    if sys.version_info >= (3, 12):
        shutil.rmtree(path, onexc=_force_remove)
    else:
        shutil.rmtree(path, onerror=_force_remove)


class FileEntry:
    """フォルダ一覧 1 行分。Explorer の「詳細表示」に相当する情報を持つ。"""

    def __init__(self, name: str, size: int, modified: datetime, is_dir: bool = False) -> None:
        self.name = name
        self.size = size
        self.modified = modified
        self.is_dir = is_dir

    @property
    def kind(self) -> str:
        if self.is_dir:
            return "ファイル フォルダー"
        ext = Path(self.name).suffix.lower()
        return {
            ".csv": "CSV ファイル",
            ".txt": "テキスト ドキュメント",
            ".log": "テキスト ドキュメント",
            ".xml": "XML ドキュメント",
            ".json": "JSON ファイル",
            ".dat": "DAT ファイル",
        }.get(ext, f"{ext.lstrip('.').upper() or 'ファイル'}")

    @property
    def size_text(self) -> str:
        if self.is_dir:
            return ""
        kb = max(1, (self.size + 1023) // 1024)
        return f"{kb:,} KB"


def list_dir(
    directory: Path,
    exclude_patterns: Optional[List[str]] = None,
    recursive: bool = False,
) -> List[FileEntry]:
    """フォルダの内容を取得する。存在しない場合は空リスト（証跡上は「フォルダ無し」を示す）。"""
    if not directory.is_dir():
        return []

    exclude_patterns = exclude_patterns or []
    paths = sorted(directory.rglob("*")) if recursive else sorted(directory.iterdir())

    entries: List[FileEntry] = []
    for p in paths:
        name = str(p.relative_to(directory)) if recursive else p.name
        if any(fnmatch.fnmatch(p.name, pat) for pat in exclude_patterns):
            continue
        try:
            st = p.stat()
        except OSError:
            continue
        entries.append(
            FileEntry(
                name=name,
                size=st.st_size,
                modified=datetime.fromtimestamp(st.st_mtime),
                is_dir=p.is_dir(),
            )
        )
    # フォルダ優先 → 名前順（Explorer の既定並び）
    entries.sort(key=lambda e: (not e.is_dir, e.name.lower()))
    return entries


def find_files(directory: Path, pattern: str = "*", recursive: bool = False) -> List[Path]:
    if not directory.is_dir():
        return []
    it = directory.rglob(pattern) if recursive else directory.glob(pattern)
    return sorted(p for p in it if p.is_file())


def clear_dir(settings: Settings, alias: str, dry_run: bool = False) -> int:
    """論理名で指定されたフォルダの中身を削除する。削除件数を返す。"""
    aliases = settings.path_aliases
    if alias not in aliases:
        raise ConfigError(
            f"クリア対象 '{alias}' は settings.paths に未定義です。"
            f"事故防止のため論理名のみ指定できます。定義済み: {sorted(aliases)}"
        )
    directory = aliases[alias]
    # 論理名チェックだけでは「その論理名が C:/ を指していた」場合を防げないので、
    # 実際に解決されたパスを削除前に必ず検証する
    assert_safe_to_clear(directory, settings.project_root)

    if not directory.is_dir():
        return 0

    count = 0
    for p in directory.iterdir():
        if dry_run:
            count += 1
            continue
        try:
            if p.is_dir():
                _rmtree(p)
            else:
                try:
                    p.unlink()
                except PermissionError:
                    os.chmod(p, stat.S_IWRITE)
                    p.unlink()
            count += 1
        except OSError as exc:
            raise ConfigError(
                f"{p} の削除に失敗しました: {exc}\n"
                f"  batch 本体・エクスプローラー・エディタ・ウイルス対策ソフトが"
                f"掴んでいないか確認してください。"
            ) from exc
    return count


def ensure_dirs(settings: Settings, aliases: List[str]) -> None:
    """batch が前提とするフォルダを作成する。"""
    for alias in aliases:
        settings.resolve_dir(alias).mkdir(parents=True, exist_ok=True)


def put_input_files(
    settings: Settings,
    case_dir: Path,
    specs: List[dict],
    dry_run: bool = False,
) -> List[Path]:
    """テストデータを投入フォルダへコピーする。コピー先のパス一覧を返す。"""
    placed: List[Path] = []
    for spec in specs:
        src_raw = spec.get("src")
        if not src_raw:
            raise ConfigError(f"setup.input_files に src がありません: {spec}")

        src = Path(src_raw)
        if not src.is_absolute():
            src = case_dir / src
        if not src.exists():
            raise ConfigError(f"投入ファイルが見つかりません: {src}")

        dest_dir = settings.resolve_dir(spec.get("dest_dir", "input_dir"))
        dest_name = spec.get("rename") or src.name
        dest = dest_dir / dest_name

        if not dry_run:
            dest_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
        placed.append(dest)
    return placed


def collect_artifacts(
    settings: Settings,
    specs: List[dict],
    save_root: Path,
) -> List[Path]:
    """処理後フォルダから成果ファイルを回収し、証跡フォルダへ保全する。

    batch を再実行すると元ファイルは消える/上書きされるため、判定より先にコピーを取る。
    """
    saved: List[Path] = []
    for spec in specs:
        alias = spec.get("dir", "output_dir")
        directory = settings.resolve_dir(alias)
        pattern = spec.get("pattern", "*")
        for src in find_files(directory, pattern, recursive=bool(spec.get("recursive", False))):
            # サブフォルダ構造を保って保全する。basename だけにすると
            # 別サブフォルダの同名ファイルが互いに上書きされ、証跡が失われる
            rel = src.relative_to(directory)
            dest = save_root / alias / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            saved.append(dest)
    return saved


def read_text_file(path: Path, encoding: str = "utf-8", fallbacks: Optional[List[str]] = None) -> str:
    """文字コードを順に試してテキストを読む。全滅時は置換読みして落とさない。"""
    for enc in [encoding, *(fallbacks or [])]:
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return path.read_text(encoding=encoding or "utf-8", errors="replace")


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
