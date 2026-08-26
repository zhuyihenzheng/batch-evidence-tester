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
from typing import Any, Dict, List, Optional

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


def clear_dir(
    settings: Settings,
    alias: str,
    dry_run: bool = False,
    exclude: Optional[List[str]] = None,
) -> int:
    """論理名で指定されたフォルダの中身を削除する。削除件数を返す。

    exclude に名前パターン（fnmatch）を渡すと、そこに一致する項目は残す。
    「Receive の下に Backup があり、Receive をクリアしたいが Backup の履歴は
    消したくない」という構成のためのもの。
    """
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

    # 他の論理名がこのフォルダの配下にある場合、そのフォルダは自動的に保護する。
    # 設定漏れで業務データ（Backup 等）を消す事故を防ぐための保険。
    protected = {settings.expand(str(p)) for p in (exclude or [])}
    for other_alias, other_path in aliases.items():
        if other_alias == alias:
            continue
        if _is_relative_to(other_path, directory) and other_path != directory:
            # directory 直下の、その論理名へ至る先頭要素を保護する
            protected.add(other_path.relative_to(directory).parts[0])

    count = 0
    for p in directory.iterdir():
        if any(fnmatch.fnmatch(p.name, pat) for pat in protected):
            continue
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


def remove_dir(settings: Settings, alias: str, dry_run: bool = False) -> bool:
    """フォルダ自体を削除する。「出力先フォルダが無い」異常系の再現に使う。

    clear_dir が中身だけ消すのに対し、こちらはフォルダごと消す。
    削除可否の判定は clear_dir と同じ安全チェックを通す。
    戻り値は実際に削除したか（元から無ければ False）。
    """
    aliases = settings.path_aliases
    if alias not in aliases:
        raise ConfigError(
            "削除対象 '%s' は settings.paths に未定義です。定義済み: %s"
            % (alias, sorted(aliases))
        )
    directory = aliases[alias]
    assert_safe_to_clear(directory, settings.project_root)

    if not directory.exists():
        return False
    if dry_run:
        return True
    try:
        _rmtree(directory)
    except OSError as exc:
        raise ConfigError(
            "%s の削除に失敗しました: %s\n"
            "  batch 本体・エクスプローラー・エディタが掴んでいないか確認してください。"
            % (directory, exc)
        ) from exc
    return True


def corrupt_bytes(data: bytes, spec: Dict[str, Any]) -> bytes:
    """正常なファイルから壊れたファイルを作る。

    解凍エラー・フォーマット不正の異常系は「壊れたファイル」を用意しないと
    再現できないが、壊れたバイナリを直接リポジトリへ置くと、中身が読めず
    「何がどう壊れているのか」が後から分からなくなる。正常な資材 1 つから
    決まった壊し方で生成すれば、意図が定義に残る。

    mode:
      truncate  … 末尾を切る（size / ratio で位置指定。既定は半分）
      empty     … 0 バイトにする
      flip      … バイト反転（既定は中央＝データ部。CRC エラーを起こす）
      garbage   … 先頭を別バイト列で上書き（マジックナンバーを壊す）
      append    … 末尾に余分なデータを足す

    どの壊し方が有効かは形式による。zip の場合:
      truncate / empty  … 中央ディレクトリごと失われるため確実に壊れる
      flip（データ部）   … 展開時に CRC エラー
      garbage（先頭）    … ローカルヘッダは壊れるが、中央ディレクトリから
                          読む実装では検出されないことがある
    実際に batch がどう振る舞うかは、まず 1 ケース試して確認すること。
    """
    mode = str(spec.get("mode", "truncate"))

    if mode == "empty":
        return b""

    if mode == "truncate":
        if "size" in spec:
            keep = max(0, int(spec["size"]))
        else:
            keep = int(len(data) * float(spec.get("ratio", 0.5)))
        return data[:keep]

    if mode == "flip":
        if not data:
            return data
        # 既定は中央付近。末尾を反転しても zip は末尾の中央ディレクトリから
        # 読むため壊れたと判定されないことがある。データ部を壊すのが確実
        offset = int(spec.get("offset", len(data) // 2))
        offset = min(max(offset, 0), len(data) - 1)
        mutated = bytearray(data)
        mutated[offset] ^= 0xFF
        return bytes(mutated)

    if mode == "garbage":
        head = str(spec.get("content", "NOT_A_VALID_FILE")).encode("utf-8")
        return head + data[len(head):] if len(data) > len(head) else head

    if mode == "append":
        extra = str(spec.get("content", "GARBAGE")).encode("utf-8")
        return data + extra

    raise ConfigError(
        "corrupt.mode が不正です: %r（truncate / empty / flip / garbage / append）" % mode)


def put_input_files(
    settings: Settings,
    case_dir: Path,
    specs: List[dict],
    dry_run: bool = False,
) -> List[Path]:
    """テストデータを投入フォルダへコピーする。コピー先のパス一覧を返す。

    src にはファイルとフォルダの両方を指定できる。フォルダの場合は
    ``dest_dir / (rename または src のフォルダ名)`` へ、配下の構造を保ったまま
    コピーする。コピー先が既にあれば、ファイル投入の上書き動作と同じく
    同名ファイルを上書きし、それ以外の既存ファイルは残す。
    """
    placed: List[Path] = []
    for spec in specs:
        src_raw = spec.get("src")
        if not src_raw:
            raise ConfigError(f"setup.input_files に src がありません: {spec}")

        src = Path(src_raw)
        if not src.is_absolute():
            src = case_dir / src
        if not src.exists():
            raise ConfigError(f"投入元が見つかりません: {src}")

        dest_dir = settings.resolve_dir(spec.get("dest_dir", "input_dir"))
        # 投入時のリネームにも {date} を使えるようにする
        dest_name = settings.expand(spec.get("rename")) if spec.get("rename") else src.name
        dest = dest_dir / dest_name

        corrupt = spec.get("corrupt")
        if corrupt and src.is_dir():
            # dry-run でも設定不備は見逃さない。実行時だけ失敗すると、事前確認の
            # 役割を果たせないため、ファイルシステムを触る前に判定する。
            raise ConfigError(
                "フォルダの投入には corrupt を指定できません: %s" % src)

        if not dry_run:
            dest_dir.mkdir(parents=True, exist_ok=True)
            if corrupt:
                # 壊れたファイルは正常な資材から生成する（意図が定義に残る）
                if not isinstance(corrupt, dict):
                    corrupt = {"mode": str(corrupt)}
                dest.write_bytes(corrupt_bytes(src.read_bytes(), corrupt))
            elif src.is_dir():
                if dest.exists() and not dest.is_dir():
                    raise ConfigError(
                        "投入先に同名のファイルがあるため、フォルダをコピーできません: %s"
                        % dest)
                # shutil.copytree(..., dirs_exist_ok=True) は Python 3.8 以降。
                # 対象環境の Python 3.6 でも同じマージ動作になるよう明示的に辿る。
                for root, dirs, files in os.walk(str(src)):
                    root_path = Path(root)
                    relative = root_path.relative_to(src)
                    target_root = dest / relative
                    target_root.mkdir(parents=True, exist_ok=True)
                    for dirname in dirs:
                        (target_root / dirname).mkdir(parents=True, exist_ok=True)
                    for filename in files:
                        target = target_root / filename
                        if target.exists() and target.is_dir():
                            raise ConfigError(
                                "投入先に同名のフォルダがあるため、ファイルをコピーできません: %s"
                                % target)
                        shutil.copy2(root_path / filename, target)
            else:
                if dest.exists() and dest.is_dir():
                    raise ConfigError(
                        "投入先に同名のフォルダがあるため、ファイルをコピーできません: %s"
                        % dest)
                shutil.copy2(src, dest)
        placed.append(dest)
    return placed


class ReplacedFile:
    """差し替えたファイルと、その退避先。実行後に元へ戻すために持ち回る。"""

    def __init__(self, dest: Path, backup: Optional[Path], existed: bool) -> None:
        self.dest = dest
        self.backup = backup      # 退避先。元ファイルが無かった場合は None
        self.existed = existed    # 差し替え前に存在したか


def replace_files(
    settings: Settings,
    case_dir: Path,
    specs: List[dict],
    backup_root: Path,
    dry_run: bool = False,
) -> List[ReplacedFile]:
    """batch の設定ファイル等をケース用のものに差し替える。

    元ファイルは backup_root へ退避する。テストが環境を変更したまま
    終わると、次のテストや手動実行が壊れた設定で動いてしまうため、
    呼び出し側は必ず restore_files() で戻すこと。
    """
    replaced: List[ReplacedFile] = []
    for spec in specs:
        src_raw = spec.get("src")
        if not src_raw:
            raise ConfigError("setup.replace_files に src がありません: %s" % spec)
        src = Path(src_raw)
        if not src.is_absolute():
            src = case_dir / src
        if not src.exists():
            raise ConfigError("差し替え元ファイルがありません: %s" % src)

        dest_dir = settings.resolve_dir(spec.get("dest_dir", ""))
        name = settings.expand(spec.get("name") or src.name)
        dest = dest_dir / name

        backup = None
        existed = dest.exists()
        if not dry_run:
            if existed:
                backup_root.mkdir(parents=True, exist_ok=True)
                backup = backup_root / (name + ".orig")
                shutil.copy2(dest, backup)
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
        replaced.append(ReplacedFile(dest, backup, existed))
    return replaced


def restore_files(replaced: List[ReplacedFile]) -> List[str]:
    """差し替えたファイルを元に戻す。戻せなかったものの説明を返す。

    1 つ失敗しても残りは戻す。環境を中途半端な状態で放置しないため。
    """
    problems: List[str] = []
    for item in replaced:
        try:
            if item.existed:
                # 元ファイルがあったなら、退避から必ず書き戻す。
                # 退避が失われていたら黙って諦めず、変更が残っている事実を報告する
                if item.backup is None or not item.backup.exists():
                    problems.append(
                        "%s の退避ファイルが見つからず復元できませんでした（差し替えたまま残っています）"
                        % item.dest)
                    continue
                shutil.copy2(item.backup, item.dest)
            elif item.dest.exists():
                # 元々無かったファイルなので、差し替えで作った分を消す
                item.dest.unlink()
        except OSError as exc:
            problems.append("%s を元に戻せませんでした: %s" % (item.dest, exc))
    return problems


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
        pattern = settings.expand(spec.get("pattern", "*"))
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
