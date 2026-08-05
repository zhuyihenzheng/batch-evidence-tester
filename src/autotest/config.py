"""設定ファイル / テストケース定義の読み込みと検証。

設計方針:
  - ケース定義はフォルダを「論理名」(input_dir 等) で参照する。物理パスは
    settings.yaml の paths だけが持ち、環境差はそこで吸収する。
  - 起動時に paths の解決と必須項目の検証を行い、実行途中で KeyError が出ないようにする。
"""

import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import yaml


class ConfigError(Exception):
    """設定不備。実行前に検出してメッセージだけ出して終了させる。"""


# 日付プレースホルダ。batch が日付ごとのフォルダ（Backup/20260803 等）を
# 使う構成に対応するためのもの。
#   {date}            -> 20260803        （業務日付。既定は本日）
#   {date:%Y/%m/%d}   -> 2026/08/03      （strftime 書式を指定）
#   {date-1}          -> 20260802        （業務日付の 1 日前）
#   {date+1:%Y%m}     -> 202609          （日数オフセット + 書式）
_DATE_PLACEHOLDER = re.compile(r"\{date(?P<offset>[+-]\d+)?(?::(?P<fmt>[^}]+))?\}")


def expand_date_placeholders(text: str, base_date: date) -> str:
    """文字列中の日付プレースホルダを展開する。

    パス・実行引数・ファイル名パターンで共通に使う。プレースホルダが
    無ければ元の文字列をそのまま返すので、無条件に通してよい。
    """
    if "{date" not in text:
        return text

    def replace(m) -> str:
        offset = int(m.group("offset") or 0)
        fmt = m.group("fmt") or "%Y%m%d"
        return (base_date + timedelta(days=offset)).strftime(fmt)

    return _DATE_PLACEHOLDER.sub(replace, text)


def _load_yaml(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise ConfigError(f"設定ファイルが見つかりません: {path}")
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        raise ConfigError(f"YAML のトップレベルはマッピングである必要があります: {path}")
    return data


class Settings:
    """settings.yaml の内容。生 dict を持ちつつ、よく使う値はプロパティで公開する。"""

    def __init__(self, raw: Dict[str, Any], source: Path, project_root: Path) -> None:
        self.raw = raw
        self.source = source
        self.project_root = project_root
        # パス等の {date} を展開するときの基準日。既定は本日。
        # --date で実行単位に、ケースの execute.date でケース単位に上書きできる。
        # ケースは直列実行されるため、ケースごとの差し替えで競合は起きない。
        self.base_date = date.today()

    def set_base_date(self, value: Union[str, date, None]) -> None:
        """基準日を設定する。文字列は YYYYMMDD または YYYY-MM-DD を受け付ける。"""
        if value is None:
            self.base_date = date.today()
            return
        if isinstance(value, datetime):
            self.base_date = value.date()
            return
        if isinstance(value, date):
            self.base_date = value
            return
        text = str(value).strip()
        for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d"):
            try:
                self.base_date = datetime.strptime(text, fmt).date()
                return
            except ValueError:
                continue
        raise ConfigError(
            "日付として解釈できません: %r（YYYYMMDD / YYYY-MM-DD 形式で指定してください）" % value
        )

    def expand(self, text: str) -> str:
        """文字列中の {date} 系プレースホルダを基準日で展開する。"""
        return expand_date_placeholders(str(text), self.base_date)

    # --- section accessors -------------------------------------------------
    @property
    def env(self) -> Dict[str, Any]:
        return self.raw.get("env", {})

    @property
    def batch(self) -> Dict[str, Any]:
        """既定の batch 定義。ケースが batch 名を指定しない場合に使われる。"""
        return self.raw.get("batch", {})

    @property
    def batches(self) -> Dict[str, Any]:
        """名前付き batch 定義。1 回の実行で複数の .exe を扱う場合に使う。"""
        return self.raw.get("batches", {}) or {}

    def batch_profile(self, name: Optional[str] = None) -> Dict[str, Any]:
        """ケースが指定した batch 名から実行設定を組み立てる。

        名前付き定義は `batch:`（既定）を土台にした差分として書ける。
        console_encoding や timeout_sec のような共通項目を毎回書かずに済む。
        """
        if not name:
            return dict(self.batch)
        if name not in self.batches:
            raise ConfigError(
                "batch '%s' は settings.yaml の batches に未定義です。定義済み: %s"
                % (name, sorted(self.batches) or "（なし）")
            )
        merged = dict(self.batch)
        merged.update(self.batches[name] or {})
        return merged

    @property
    def database(self) -> Dict[str, Any]:
        return self.raw.get("database", {})

    @property
    def log(self) -> Dict[str, Any]:
        return self.raw.get("log", {})

    @property
    def evidence(self) -> Dict[str, Any]:
        return self.raw.get("evidence", {})

    @property
    def excel(self) -> Dict[str, Any]:
        return self.raw.get("excel", {})

    @property
    def folder_evidence(self) -> Dict[str, Any]:
        return self.raw.get("folder_evidence", {})

    @property
    def tester(self) -> str:
        return self.env.get("tester") or os.environ.get("USERNAME") or os.environ.get("USER") or "unknown"

    # --- path resolution ---------------------------------------------------
    @property
    def path_aliases(self) -> Dict[str, Path]:
        """論理名 -> 物理パス（絶対パスに正規化済み）。"""
        return {name: self._absolute(str(p)) for name, p in (self.raw.get("paths") or {}).items()}

    def resolve_dir(self, alias_or_path: str) -> Path:
        """論理名なら settings.paths を引く。それ以外は生パスとして扱う。"""
        aliases = self.path_aliases
        if alias_or_path in aliases:
            return aliases[alias_or_path]
        return self._absolute(alias_or_path)

    def _absolute(self, raw: str) -> Path:
        """相対パスはプロジェクトルート基準にする。

        カレントフォルダに依存すると、タスクスケジューラ起動時と手動起動時で
        参照先が変わってしまうため。{date} 等のプレースホルダもここで展開する。
        """
        path = Path(self.expand(raw))
        return path if path.is_absolute() else (self.project_root / path).resolve()

    def db_password(self) -> str:
        env_name = self.database.get("password_env")
        if not env_name:
            return ""
        pw = os.environ.get(env_name)
        if pw is None:
            raise ConfigError(
                f"DB パスワードの環境変数 {env_name} が未設定です。\n"
                f"  Windows: setx {env_name} \"<password>\"  （設定後にコンソールを開き直す）\n"
                f"  macOS/Linux: export {env_name}='<password>'"
            )
        return pw


def load_settings(path: Union[str, Path], project_root: Union[str, Optional[Path]] = None) -> Settings:
    path = Path(path)
    root = Path(project_root) if project_root else path.resolve().parent.parent
    settings = Settings(raw=_load_yaml(path), source=path, project_root=root)
    _validate_settings(settings)
    return settings


def _validate_settings(s: Settings) -> None:
    missing: List[str] = []
    if not s.batch.get("exe_path"):
        missing.append("batch.exe_path")
    if not s.database.get("server"):
        missing.append("database.server")
    if not s.database.get("database"):
        missing.append("database.database")
    if not s.raw.get("paths"):
        missing.append("paths")
    if missing:
        raise ConfigError(f"{s.source} に必須項目がありません: {', '.join(missing)}")

    # folder_evidence.targets が paths に存在するか
    aliases = set(s.path_aliases)
    unknown = [t for t in s.folder_evidence.get("targets", []) if t not in aliases]
    if unknown:
        raise ConfigError(
            f"folder_evidence.targets に paths 未定義の論理名があります: {unknown}\n"
            f"  定義済み: {sorted(aliases)}"
        )


# =============================================================================
# テストケース
# =============================================================================


class TestCase:
    """cases/*.yaml 1 ファイル分。"""

    def __init__(
        self,
        case_id: str,
        name: str,
        source: Path,
        description: str = "",
        tags: Optional[List[str]] = None,
        enabled: bool = True,
        setup: Optional[Dict[str, Any]] = None,
        snapshot: Optional[Dict[str, Any]] = None,
        execute: Optional[Dict[str, Any]] = None,
        collect: Optional[Dict[str, Any]] = None,
        assertions: Optional[Dict[str, Any]] = None,
        teardown: Optional[Dict[str, Any]] = None,
    ) -> None:
        self.case_id = case_id
        self.name = name
        self.source = source
        self.description = description
        self.tags = tags if tags is not None else []
        self.enabled = enabled
        self.setup = setup if setup is not None else {}
        self.snapshot = snapshot if snapshot is not None else {}
        self.execute = execute if execute is not None else {}
        self.collect = collect if collect is not None else {}
        self.assertions = assertions if assertions is not None else {}
        self.teardown = teardown if teardown is not None else {}

    @property
    def dir(self) -> Path:
        """ケース固有の資材（投入ファイル等）を置くフォルダ。"""
        return self.source.parent / self.case_id if (self.source.parent / self.case_id).is_dir() else self.source.parent


def _folder_tags(path: Path, cases_dir: Path) -> List[str]:
    """ケース定義の置き場所から暗黙のタグを作る。

    cases/受注/TC001.yaml       -> ["受注"]
    cases/受注/取込/TC001.yaml   -> ["受注", "取込"]
    cases/TC001.yaml            -> []
    """
    try:
        rel = path.relative_to(cases_dir)
    except ValueError:
        return []
    return [part for part in rel.parts[:-1]]


def find_case_files(cases_dir: Path) -> List[Path]:
    """ケース定義 YAML を再帰的に集める。

    機能ごとにサブフォルダで整理できるようにする:
        cases/受注/TC001.yaml
        cases/請求/TC010.yaml

    ケース資材（cases/<機能>/TC001/input/... 等）の中にある YAML は
    ケース定義ではないので除外する。判定基準は「同名の YAML が兄弟に
    存在するフォルダの配下かどうか」。
    """
    all_yaml = sorted(list(cases_dir.rglob("*.yaml")) + list(cases_dir.rglob("*.yml")))
    # 資材フォルダ = 同名の定義ファイルが隣にあるフォルダ
    material_dirs = {
        p.parent / p.stem
        for p in all_yaml
        if (p.parent / p.stem).is_dir()
    }

    def is_material(path: Path) -> bool:
        for parent in path.parents:
            if parent in material_dirs:
                return True
        return False

    return [p for p in all_yaml if not is_material(p)]


def load_cases(cases_dir: Union[str, Path], only: Optional[List[str]] = None, tags: Optional[List[str]] = None) -> List[TestCase]:
    """cases/ 配下の *.yaml を読み込む。only / tags で絞り込み可能。

    サブフォルダ名は暗黙のタグになる（cases/受注/TC001.yaml なら「受注」）。
    機能ごとにフォルダを切っておけば --tag 受注 でまとめて実行できる。
    """
    cases_dir = Path(cases_dir)
    if not cases_dir.is_dir():
        raise ConfigError(f"ケースフォルダがありません: {cases_dir}")

    cases: List[TestCase] = []
    for path in find_case_files(cases_dir):
        data = _load_yaml(path)
        case_id = str(data.get("id") or path.stem)
        case = TestCase(
            case_id=case_id,
            name=str(data.get("name") or case_id),
            source=path,
            description=str(data.get("description") or ""),
            # サブフォルダ名を暗黙のタグとして足す。機能別にフォルダを切れば
            # そのまま --tag <機能名> で絞り込める（宣言済みのタグは重複させない）
            tags=_folder_tags(path, cases_dir) + [
                t for t in (data.get("tags") or []) if t not in _folder_tags(path, cases_dir)
            ],
            enabled=bool(data.get("enabled", True)),
            setup=data.get("setup") or {},
            snapshot=data.get("snapshot") or {},
            execute=data.get("execute") or {},
            collect=data.get("collect") or {},
            # "assert" は Python の予約語なので属性名を変える
            assertions=data.get("assert") or {},
            teardown=data.get("teardown") or {},
        )
        cases.append(case)

    if not cases:
        raise ConfigError(f"{cases_dir} にテストケース (*.yaml) がありません")

    # ID 重複は証跡フォルダと Excel シートが上書きし合うため設定エラーにする
    seen: Dict[str, Path] = {}
    for c in cases:
        if c.case_id in seen:
            raise ConfigError(
                f"ケース ID '{c.case_id}' が重複しています:\n"
                f"  {seen[c.case_id]}\n  {c.source}\n"
                f"  証跡フォルダと Excel シートが上書きし合うため、ID は一意にしてください。"
            )
        seen[c.case_id] = c.source

    if only:
        wanted = set(only)
        found = {c.case_id for c in cases}
        unknown = wanted - found
        if unknown:
            raise ConfigError(f"指定されたケース ID が存在しません: {sorted(unknown)} / 定義済み: {sorted(found)}")
        cases = [c for c in cases if c.case_id in wanted]

    if tags:
        want_tags = set(tags)
        cases = [c for c in cases if want_tags & set(c.tags)]
        if not cases:
            raise ConfigError(f"タグ {tags} に一致するケースがありません")

    enabled_cases = [c for c in cases if c.enabled]
    if not enabled_cases:
        # 0 ケース実行で正常終了すると「何もテストしていないのに緑」になるため拒否する
        disabled = [c.case_id for c in cases if not c.enabled]
        raise ConfigError(
            f"実行可能なケースが 0 件です（enabled: false のケース: {disabled}）。\n"
            f"  少なくとも 1 件を enabled にするか、絞り込み条件を見直してください。"
        )
    return enabled_cases
