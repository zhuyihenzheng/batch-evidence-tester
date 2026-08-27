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
#   {date}            -> 20260803        （基準日。既定は本日）
#   {date:%Y/%m/%d}   -> 2026/08/03      （strftime 書式を指定）
#   {date-1}          -> 20260802        （基準日の 1 日前）
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
        review: Optional[List[str]] = None,
        assertions: Optional[Dict[str, Any]] = None,
        teardown: Optional[Dict[str, Any]] = None,
        mode: str = "auto",
    ) -> None:
        self.case_id = case_id
        self.name = name
        self.source = source
        self.description = description
        self.tags = tags if tags is not None else []
        self.enabled = enabled
        self.mode = mode
        self.setup = setup if setup is not None else {}
        self.snapshot = snapshot if snapshot is not None else {}
        self.execute = execute if execute is not None else {}
        self.collect = collect if collect is not None else {}
        # 証跡を見て人が最終確認する観点。自動比較の assert とは分離する。
        self.review = review if review is not None else []
        self.assertions = assertions if assertions is not None else {}
        self.teardown = teardown if teardown is not None else {}

    @property
    def dir(self) -> Path:
        """ケース固有の資材（投入ファイル等）を置くフォルダ。"""
        return self.source.parent / self.case_id if (self.source.parent / self.case_id).is_dir() else self.source.parent

    @property
    def is_manual(self) -> bool:
        """人が手で batch を動かすケースか。autotest run の対象外になる。"""
        return self.mode == "manual"


# ケース定義で使えるトップレベル項目。ここに無いキーは綴り間違いとして扱う。
# 黙って無視すると「書いたのに効かない」状態が静かに続く。
CASE_TOP_KEYS = {
    "id", "name", "description", "tags", "enabled", "mode",
    "setup", "snapshot", "execute", "collect", "review", "assert", "teardown",
}
# 実行方式。
#   auto   … autotest run が .exe を起動して自動判定する（既定）
#   manual … 人が手で batch を動かす。run からは除外され、
#            autotest manual --phase before / after で証跡だけ採る
CASE_MODES = ("auto", "manual")
SETUP_KEYS = {
    "clean_dirs", "remove_dirs", "sql", "input_files", "replace_files", "batches",
    "sql_after_batches", "db_lock",
}
SETUP_BATCH_KEYS = {"batch", "args", "expected_exit_code"}
COLLECT_KEYS = {"files", "folder_evidence"}
EXECUTE_KEYS = {"batch", "args", "date", "expected_exit_code"}
ASSERT_KEYS = {"exit_code", "db", "files", "log"}
DB_ASSERT_KEYS = {"table", "expected", "key", "ignore_columns", "manual"}
FILE_ASSERT_KEYS = {"name", "actual", "expected", "encoding", "ignore_line_patterns",
                    "exists", "manual"}
FILE_ACTUAL_KEYS = {"dir", "pattern"}
FILE_EXISTS_KEYS = {"dir", "pattern", "count"}
LOG_ASSERT_KEYS = {"must_contain", "must_not_contain", "manual"}
EXIT_CODE_ASSERT_KEYS = {"expected", "manual"}


def _as_bool(value: Any, where: str) -> bool:
    """YAML の真偽値を厳密に解釈する。

    enabled: "false" は文字列なので Python では真になる。これを黙って
    True と扱うと、無効化したつもりのケースが実行されてしまう。
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in ("true", "yes", "on", "1"):
            return True
        if lowered in ("false", "no", "off", "0"):
            return False
    raise ConfigError(
        "%s は true / false で指定してください（指定値: %r）。"
        "引用符で囲むと文字列になり、意図と逆に解釈されます。" % (where, value)
    )


def _check_unknown_keys(data: Dict[str, Any], allowed: set, where: str, path: Path) -> List[str]:
    unknown = [k for k in data if k not in allowed]
    if not unknown:
        return []
    return ["%s: 未知の項目 %s（綴り間違いの可能性。使える項目: %s）"
            % (where, sorted(unknown), sorted(allowed))]


def _check_manual(value: Any, where: str) -> List[str]:
    """manual は厳密な YAML boolean だけを受け付ける。

    bool("false") は True になるため、文字列を許すと利用者の指定と逆の動作に
    なり得る。綴り間違いと同じく、実行前に設定エラーとして止める。
    """
    if value is None or isinstance(value, bool):
        return []
    return ["%s は true / false で指定してください（引用符は付けないでください。指定値: %r）"
            % (where, value)]


def _validate_assertions(assertions: Dict[str, Any], path: Path) -> List[str]:
    """assert 配下を項目単位で検証し、書いたのに効かない設定を作らせない。"""
    problems: List[str] = []

    exit_code = assertions.get("exit_code")
    if isinstance(exit_code, dict):
        problems += _check_unknown_keys(
            exit_code, EXIT_CODE_ASSERT_KEYS, "assert.exit_code", path)
        problems += _check_manual(exit_code.get("manual"), "assert.exit_code.manual")
        if "expected" not in exit_code:
            problems.append("assert.exit_code に expected がありません")
        elif not isinstance(exit_code.get("expected"), int) or isinstance(exit_code.get("expected"), bool):
            problems.append("assert.exit_code.expected は整数で指定してください")
    elif exit_code is not None and (not isinstance(exit_code, int) or isinstance(exit_code, bool)):
        problems.append("assert.exit_code は整数、または {expected: 整数, manual: true} で指定してください")

    db_items = assertions.get("db", [])
    if db_items is None:
        db_items = []
    if not isinstance(db_items, list):
        problems.append("assert.db はリストで指定してください")
    else:
        for i, item in enumerate(db_items):
            where = "assert.db[%d]" % i
            if not isinstance(item, dict):
                problems.append("%s はマッピングで指定してください: %r" % (where, item))
                continue
            problems += _check_unknown_keys(item, DB_ASSERT_KEYS, where, path)
            problems += _check_manual(item.get("manual"), where + ".manual")
            if not item.get("table"):
                problems.append("%s に table がありません" % where)
            if not item.get("expected") and item.get("manual") is not True:
                problems.append("%s に expected がありません（目視確認なら manual: true を指定してください）"
                                % where)

    file_items = assertions.get("files", [])
    if file_items is None:
        file_items = []
    if not isinstance(file_items, list):
        problems.append("assert.files はリストで指定してください")
    else:
        for i, item in enumerate(file_items):
            where = "assert.files[%d]" % i
            if not isinstance(item, dict):
                problems.append("%s はマッピングで指定してください: %r" % (where, item))
                continue
            problems += _check_unknown_keys(item, FILE_ASSERT_KEYS, where, path)
            problems += _check_manual(item.get("manual"), where + ".manual")
            has_actual = "actual" in item
            has_exists = "exists" in item
            if has_actual == has_exists:
                problems.append("%s は actual または exists のどちらか一方を指定してください" % where)
                continue
            location_key = "actual" if has_actual else "exists"
            location = item.get(location_key)
            if not isinstance(location, dict):
                problems.append("%s.%s はマッピングで指定してください" % (where, location_key))
                continue
            problems += _check_unknown_keys(
                location, FILE_ACTUAL_KEYS if has_actual else FILE_EXISTS_KEYS,
                "%s.%s" % (where, location_key), path)
            if has_actual and not item.get("expected") and item.get("manual") is not True:
                problems.append("%s に expected がありません（目視確認なら manual: true を指定してください）"
                                % where)
            if has_exists and any(key in item for key in ("expected", "encoding", "ignore_line_patterns")):
                problems.append("%s の exists 判定では expected / encoding / ignore_line_patterns は使えません"
                                % where)

    log_spec = assertions.get("log")
    if log_spec is not None:
        if not isinstance(log_spec, dict):
            problems.append("assert.log はマッピングで指定してください")
        else:
            problems += _check_unknown_keys(log_spec, LOG_ASSERT_KEYS, "assert.log", path)
            problems += _check_manual(log_spec.get("manual"), "assert.log.manual")

    return problems


def _validate_review(review: Any, path: Path) -> List[str]:
    """人工確認内容の文字列リストを検証する。"""
    if review is None:
        return []
    if not isinstance(review, list):
        return ["review はリストで指定してください"]

    problems: List[str] = []
    contents = set()
    for i, item in enumerate(review):
        where = "review[%d]" % i
        if not isinstance(item, str) or not item.strip():
            problems.append("%s は空でない確認内容の文字列で指定してください" % where)
            continue
        normalized = item.strip()
        if normalized in contents:
            problems.append("review の確認内容が重複しています: %s" % normalized)
        contents.add(normalized)
    return problems


def validate_case_id(case_id: Any) -> List[str]:
    """ケース ID として使える文字列かを検証する。問題の一覧を返す（空なら合格）。

    ID は証跡フォルダ名と Excel シート名にそのまま使うため、パス区切りや
    Windows で使えない記号を弾く。ケース定義の検証と scaffold（新規作成・複製）
    で同じ規則を使うために切り出してある。
    """
    problems: List[str] = []
    text = str(case_id)
    if not text.strip():
        problems.append("id が空です")
        return problems
    if any(ch in text for ch in "/\\:*?\"<>|") or text in (".", ".."):
        problems.append("id にパス区切りや記号は使えません: %r" % case_id)
    return problems


def _validate_case_schema(data: Dict[str, Any], path: Path, case_id: str) -> List[str]:
    """ケース定義の型と項目名を検証する。実行時に KeyError で落ちるのを防ぐ。"""
    problems: List[str] = []
    problems += _check_unknown_keys(data, CASE_TOP_KEYS, "トップレベル", path)
    problems += validate_case_id(case_id)

    if "enabled" in data:
        try:
            _as_bool(data["enabled"], "enabled")
        except ConfigError as exc:
            problems.append(str(exc))

    tags = data.get("tags")
    if tags is not None and not isinstance(tags, list):
        problems.append("tags はリストで指定してください（指定値: %r）" % tags)

    if "mode" in data and data["mode"] not in CASE_MODES:
        problems.append(
            "mode は %s のいずれかで指定してください（指定値: %r）"
            % (" / ".join(CASE_MODES), data["mode"]))

    setup = data.get("setup") or {}
    if not isinstance(setup, dict):
        problems.append("setup はマッピングで指定してください")
        setup = {}
    problems += _check_unknown_keys(setup, SETUP_KEYS, "setup", path)

    # batch 実行中だけ別接続で保持するロック（DB 更新失敗の再現用）。
    # リストや辞書を渡されても orchestrator は文字列として扱うため、先に弾く
    if "db_lock" in setup and not isinstance(setup["db_lock"], str):
        problems.append("setup.db_lock は SQL 文字列で指定してください（指定値: %r）" % setup["db_lock"])

    for key in ("sql", "sql_after_batches"):
        if key in setup and not isinstance(setup[key], list):
            problems.append("setup.%s はリストで指定してください" % key)

    for key in ("clean_dirs", "remove_dirs"):
        for i, spec in enumerate(setup.get(key) or []):
            if isinstance(spec, dict):
                if not spec.get("alias"):
                    problems.append("setup.%s[%d] に alias がありません: %r" % (key, i, spec))
            elif not isinstance(spec, str):
                problems.append("setup.%s[%d] は論理名の文字列か {alias: ...} で指定してください: %r"
                                % (key, i, spec))

    for key, required in (("input_files", "src"), ("replace_files", "src")):
        for i, spec in enumerate(setup.get(key) or []):
            if not isinstance(spec, dict):
                problems.append("setup.%s[%d] はマッピングで指定してください: %r" % (key, i, spec))
            elif not spec.get(required):
                problems.append("setup.%s[%d] に %s がありません" % (key, i, required))
            else:
                for field in ("src", "rename", "name"):
                    problems += _check_relative(spec.get(field), "setup.%s[%d].%s" % (key, i, field))

    # 投入ファイルと設定差し替えが完了した後、主 batch より前に実行する batch。
    # 同じ名前を複数回並べる用途もあるため、mapping の list として保持する。
    setup_batches = setup.get("batches")
    if setup_batches is not None and not isinstance(setup_batches, list):
        problems.append("setup.batches はリストで指定してください")
    elif isinstance(setup_batches, list):
        for i, spec in enumerate(setup_batches):
            where = "setup.batches[%d]" % i
            if isinstance(spec, str):
                # 引数無しなら `- batch_name` の短縮形も許可する。
                if not spec.strip():
                    problems.append("%s の batch 名が空です" % where)
                continue
            if not isinstance(spec, dict):
                problems.append(
                    "%s は batch 名の文字列か {batch: ..., args: [...]} で指定してください: %r"
                    % (where, spec))
                continue
            problems += _check_unknown_keys(spec, SETUP_BATCH_KEYS, where, path)
            batch_name = spec.get("batch")
            if batch_name is not None and (not isinstance(batch_name, str) or not batch_name.strip()):
                problems.append("%s.batch は空でない文字列で指定してください" % where)
            batch_args = spec.get("args")
            if batch_args is not None and not isinstance(batch_args, list):
                problems.append("%s.args はリストで指定してください" % where)
            expected = spec.get("expected_exit_code")
            if expected is not None and (not isinstance(expected, int) or isinstance(expected, bool)):
                problems.append("%s.expected_exit_code は整数で指定してください" % where)

    execute = data.get("execute") or {}
    if not isinstance(execute, dict):
        problems.append("execute はマッピングで指定してください")
        execute = {}
    problems += _check_unknown_keys(execute, EXECUTE_KEYS, "execute", path)
    args = execute.get("args")
    if args is not None and not isinstance(args, list):
        # 文字列を渡すと 1 文字ずつに分解されて .exe へ渡ってしまう
        problems.append('execute.args はリストで指定してください（例: ["--mode", "daily"]）。'
                        "文字列を渡すと 1 文字ずつ分解されます。指定値: %r" % args)

    collect = data.get("collect") or {}
    if isinstance(collect, dict):
        problems += _check_unknown_keys(collect, COLLECT_KEYS, "collect", path)
    else:
        problems.append("collect はマッピングで指定してください")

    problems += _validate_review(data.get("review"), path)

    assertions = data.get("assert")
    if assertions is None:
        assertions = {}
    if isinstance(assertions, dict):
        problems += _check_unknown_keys(assertions, ASSERT_KEYS, "assert", path)
        problems += _validate_assertions(assertions, path)
    else:
        problems.append("assert はマッピングで指定してください")

    return problems


def _check_relative(value: Any, where: str) -> List[str]:
    """資材パスが上位ディレクトリへ抜けないことを確認する。"""
    if not value or not isinstance(value, str):
        return []
    normalized = value.replace("\\", "/")
    if normalized.startswith("/") or ".." in normalized.split("/"):
        return ["%s に上位フォルダへ抜けるパスは指定できません: %r" % (where, value)]
    return []


def _folder_tags(path: Path, cases_dir: Path) -> List[str]:
    """ケース定義の置き場所から暗黙のタグを作る。

    cases/group_a/TC001.yaml              -> ["group_a"]
    cases/group_a/subgroup_1/TC001.yaml   -> ["group_a", "subgroup_1"]
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
        cases/group_a/TC001.yaml
        cases/group_b/TC010.yaml

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

    サブフォルダ名は暗黙のタグになる（cases/group_a/TC001.yaml なら「group_a」）。
    グループごとにフォルダを切っておけば --tag group_a でまとめて実行できる。
    """
    cases_dir = Path(cases_dir)
    if not cases_dir.is_dir():
        raise ConfigError(f"ケースフォルダがありません: {cases_dir}")

    cases: List[TestCase] = []
    for path in find_case_files(cases_dir):
        data = _load_yaml(path)
        case_id = str(data.get("id") or path.stem)

        schema_problems = _validate_case_schema(data, path, case_id)
        if schema_problems:
            raise ConfigError(
                "ケース定義に問題があります: %s\n  - %s" % (path, "\n  - ".join(schema_problems)))

        cases.append(_build_case(data, path, case_id, cases_dir))

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


def _build_case(data: Dict[str, Any], path: Path, case_id: str, cases_dir: Path) -> TestCase:
    """検証済みの生 dict から TestCase を組み立てる。"""
    folder_tags = _folder_tags(path, cases_dir)
    return TestCase(
        case_id=case_id,
        name=str(data.get("name") or case_id),
        source=path,
        description=str(data.get("description") or ""),
        # サブフォルダ名を暗黙のタグとして足す。機能別にフォルダを切れば
        # そのまま --tag <機能名> で絞り込める（宣言済みのタグは重複させない）
        tags=folder_tags + [t for t in (data.get("tags") or []) if t not in folder_tags],
        enabled=_as_bool(data.get("enabled", True), "enabled"),
        mode=str(data.get("mode") or "auto"),
        setup=data.get("setup") or {},
        snapshot=data.get("snapshot") or {},
        execute=data.get("execute") or {},
        collect=data.get("collect") or {},
        review=data.get("review") or [],
        # "assert" は Python の予約語なので属性名を変える
        assertions=data.get("assert") or {},
        teardown=data.get("teardown") or {},
    )
