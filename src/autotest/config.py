"""設定ファイル / テストケース定義の読み込みと検証。

設計方針:
  - ケース定義はフォルダを「論理名」(input_dir 等) で参照する。物理パスは
    settings.yaml の paths だけが持ち、環境差はそこで吸収する。
  - 起動時に paths の解決と必須項目の検証を行い、実行途中で KeyError が出ないようにする。
"""

import os
import re
from copy import deepcopy
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import yaml


class ConfigError(Exception):
    """設定不備。実行前に検出してメッセージだけ出して終了させる。"""


SETTINGS_NAMES = ("settings.yaml", "settings.yml")
LOCAL_SETTINGS_NAMES = (
    "settings.local.yaml", "settings.local.yml", "settings_local.yaml",
    "settings_local.yml", "setting_local.yaml", "setting_local.yml",
)
FOLDER_SETTINGS_KEYS = {"batch", "batches", "paths", "log", "folder_evidence", "evidence"}

# batch の起動方式。
#   exe  … exe_path をそのまま起動する（C# 製 .exe など。既定）
#   java … java [jvm_args] -jar <jar> [common_args] [execute.args] を組み立てて起動する
#          （Spring Boot の実行可能 jar を想定）
BATCH_TYPES = ("exe", "java")
# batch / batches.<名前> に書ける項目。綴り間違い（jvm_arg 等）を黙って無視すると
# 「指定したつもりの JVM オプションが効いていない」状態で試験が進むため、未知の項目は止める。
BATCH_PROFILE_KEYS = {
    "type", "exe_path", "java", "jar", "jvm_args", "working_dir", "common_args", "env_vars",
    "console_encoding", "timeout_sec", "heartbeat_sec",
    "log_settle_sec", "log_settle_min_sec", "log_settle_stable_sec", "log_dir", "folder_evidence",
}

# DB の種類。
#   sqlserver  … pyodbc + Microsoft ODBC Driver（既定）
#   postgresql … 標準ライブラリだけの内蔵クライアント（pgwire.py）。追加のインストールは不要
DB_TYPES = ("sqlserver", "postgresql")
_DATABASE_COMMON_KEYS = {
    "type", "server", "database", "user", "password_env",
    "login_timeout_sec", "query_timeout_sec", "batch_start_margin_sec",
}
# 種類ごとに書ける項目。もう一方の種類の項目は黙って無視せずに止める。
# たとえば PostgreSQL で encrypt: true を書いても暗号化は効かない（効くのは sslmode）ため、
# 「暗号化したつもり」のまま平文で繋がる
DATABASE_SQLSERVER_KEYS = _DATABASE_COMMON_KEYS | {"driver", "auth", "encrypt", "trust_server_certificate"}
DATABASE_POSTGRESQL_KEYS = _DATABASE_COMMON_KEYS | {"port", "sslmode", "sslrootcert", "timezone", "search_path"}
PG_SSLMODES = ("disable", "prefer", "require", "verify-ca", "verify-full")
PG_DEFAULT_PORT = 5432
_DATABASE_KEY_HINTS = {
    ("postgresql", "driver"): "PostgreSQL は内蔵クライアントで接続するため ODBC ドライバは使いません",
    ("postgresql", "auth"): "認証方式はサーバの pg_hba.conf で決まります。パスワードは password_env で渡します",
    ("postgresql", "encrypt"): "暗号化は sslmode で指定します（require / verify-ca / verify-full）",
    ("postgresql", "trust_server_certificate"): "証明書の検証は sslmode: verify-ca / verify-full と sslrootcert で指定します",
    ("sqlserver", "port"): "SQL Server のポートは server に host,1433 の形で書きます",
    ("sqlserver", "sslmode"): "SQL Server の暗号化は encrypt / trust_server_certificate で指定します",
    ("sqlserver", "sslrootcert"): "SQL Server の証明書は trust_server_certificate で扱います",
    ("sqlserver", "timezone"): "timezone は PostgreSQL 用の項目です",
    ("sqlserver", "search_path"): "search_path は PostgreSQL 用の項目です",
}


def _settings_file(directory: Path, names) -> Optional[Path]:
    found = [directory / name for name in names if (directory / name).is_file()]
    if len(found) > 1:
        raise ConfigError("同じ階層に設定ファイルが複数あります。1 つに統一してください: %s" % found)
    return found[0] if found else None


def default_config(project_root: Path) -> Path:
    directory = project_root / "config"
    return (_settings_file(directory, LOCAL_SETTINGS_NAMES)
            or _settings_file(directory, SETTINGS_NAMES) or directory / "settings.yaml")


def _merge_settings(base, override):
    """辞書は再帰マージ、リスト・スカラーは置換。元の設定を変更しない。"""
    merged = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _merge_settings(merged[key], value)
        else:
            merged[key] = deepcopy(value)
    return merged


def _folder_evidence_options(value, where):
    """リストの短縮形と詳細設定を同じ形式に正規化する。空リストは撮影無し。"""
    if isinstance(value, list):
        value = {"targets": value}
    if not isinstance(value, dict):
        raise ConfigError("%s はフォルダ名のリスト、またはマッピングで指定してください" % where)
    allowed = {"targets", "recursive", "exclude_patterns", "max_entries"}
    unknown = set(value) - allowed
    if unknown:
        raise ConfigError("%s に未知の項目があります: %s" % (where, sorted(unknown)))
    for key in ("targets", "exclude_patterns"):
        if key in value and (not isinstance(value[key], list) or
                             any(not isinstance(item, str) or not item.strip() for item in value[key])):
            raise ConfigError("%s.%s は空でない文字列のリストで指定してください" % (where, key))
    if "recursive" in value and not isinstance(value["recursive"], bool):
        raise ConfigError("%s.recursive は true / false で指定してください" % where)
    if "max_entries" in value:
        count = value["max_entries"]
        if not isinstance(count, int) or isinstance(count, bool) or count < 1:
            raise ConfigError("%s.max_entries は 1 以上の整数で指定してください" % where)
    return value


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


class DuplicateKeyError(yaml.constructor.ConstructorError):
    """同じマッピングに同じキーが 2 回書かれている。"""

    def __init__(self, key: Any, first_mark: Any, mark: Any) -> None:
        super(DuplicateKeyError, self).__init__(
            None, None, "キー %r が 2 回書かれています（1 回目は %d 行目）" % (key, first_mark.line + 1), mark)
        self.key = key
        self.first_mark = first_mark


class UniqueKeyLoader(yaml.SafeLoader):
    """同じキーを 2 回書いたらエラーにする SafeLoader。

    PyYAML の safe_load は後の値で黙って上書きする。assert に log を 2 回書くと 1 回目の
    検査が消えるのに、残ったデータはスキーマとして正しいので検証では見つけられない
    （AI がケースに追記するときに起きやすい、検査が減る方向の誤り）。
    """

    def construct_mapping(self, node, deep=False):
        if isinstance(node, yaml.MappingNode):
            seen = {}  # type: Dict[Any, Any]
            for key_node, _value_node in node.value:
                if key_node.tag == "tag:yaml.org,2002:merge":
                    continue   # << で取り込むキーは上書きされる前提の書き方なので対象外
                key = self.construct_object(key_node, deep=True)
                try:
                    first = seen.get(key)
                except TypeError:
                    continue   # ハッシュできないキーは親クラスがエラーにする
                if first is not None:
                    raise DuplicateKeyError(key, first, key_node.start_mark)
                seen[key] = key_node.start_mark
        return super(UniqueKeyLoader, self).construct_mapping(node, deep=deep)


def load_yaml_text(text: Any) -> Any:
    """重複キーをエラーにする safe_load。ケース・設定・md の YAML はすべてこれで読む。"""
    return yaml.load(text, Loader=UniqueKeyLoader)


def _load_yaml(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise ConfigError(f"設定ファイルが見つかりません: {path}")
    try:
        with path.open("r", encoding="utf-8") as f:
            data = load_yaml_text(f)
    except (OSError, yaml.YAMLError) as exc:
        raise ConfigError("YAML を読み込めません: %s\n%s" % (path, exc))
    if not isinstance(data, dict):
        raise ConfigError(f"YAML のトップレベルはマッピングである必要があります: {path}")
    return data


class Settings:
    """settings.yaml の内容。生 dict を持ちつつ、よく使う値はプロパティで公開する。"""

    def __init__(self, raw: Dict[str, Any], source: Path, project_root: Path) -> None:
        self.raw = raw
        self.source = source
        self.project_root = project_root
        self.sources = [source]
        self._case_source = None
        # パス等の {date} を展開するときの基準日。既定は本日。
        # --date で実行単位に、ケースの execute.date でケース単位に上書きできる。
        # ケースは直列実行されるため、ケースごとの差し替えで競合は起きない。
        self.base_date = date.today()

    def for_case(self, case: "TestCase") -> "Settings":
        """cases ルートからケースの親までの設定を、親→子の順で重ねる。"""
        case_source = case.source.resolve()
        if self._case_source == case_source:
            return self
        root = (self.project_root / "cases").resolve()
        try:
            case_source.relative_to(root)
        except ValueError:
            root = (case.cases_root or case.source.parent).resolve()
        try:
            relative = case_source.parent.relative_to(root)
        except ValueError:
            raise ConfigError("ケースが設定探索ルートの外にあります: %s" % case.source)
        directories = [root]
        for part in relative.parts:
            directories.append(directories[-1] / part)
        raw = self.raw
        sources = list(self.sources)
        for directory in directories:
            for names in (SETTINGS_NAMES, LOCAL_SETTINGS_NAMES):
                path = _settings_file(directory, names)
                if path is None or path.resolve() == self.source.resolve():
                    continue
                override = _load_yaml(path)
                unknown = set(override) - FOLDER_SETTINGS_KEYS
                if unknown:
                    raise ConfigError(
                        "%s: フォルダ設定で使えない項目: %s。env / database / excel は共通設定に置いてください。"
                        % (path, sorted(unknown)))
                raw = _merge_settings(raw, override)
                sources.append(path)
        if len(sources) == len(self.sources):
            return self
        effective = Settings(raw, self.source, self.project_root)
        effective.sources = sources
        effective._case_source = case_source
        effective.base_date = self.base_date
        _validate_settings(effective)
        return effective

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
    def db_type(self) -> str:
        return str(self.database.get("type") or "sqlserver")

    def db_label(self) -> str:
        """証跡や画面に出す接続先サーバの表記。SQL Server は従来どおり server の値そのまま。"""
        if self.db_type == "postgresql":
            host, port = pg_host_port(self.database)
            return "PostgreSQL %s:%d" % (host, port)
        return str(self.database.get("server", ""))

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

    def folder_evidence_for(self, case: "TestCase") -> Dict[str, Any]:
        """撮影・フォルダ準備・preflight が同じ実効設定を使用する。"""
        cfg = deepcopy(_folder_evidence_options(self.folder_evidence, "folder_evidence"))
        profile = self.batch_profile(case.execute.get("batch"))
        if "folder_evidence" in profile:
            cfg.update(_folder_evidence_options(profile["folder_evidence"], "batch.folder_evidence"))
        if "folder_evidence" in case.collect:
            cfg.update(_folder_evidence_options(case.collect["folder_evidence"], "collect.folder_evidence"))
        targets = cfg.get("targets", [])
        aliases = self.path_aliases
        unknown = [alias for alias in targets if alias not in aliases]
        if unknown:
            raise ConfigError("%s: folder_evidence に paths 未定義の論理名があります: %s"
                              % (case.case_id, unknown))
        return deepcopy(cfg)

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


def load_settings(path: Union[str, Path], project_root: Union[str, Optional[Path]] = None,
                  require_runtime: bool = True) -> Settings:
    path = Path(path)
    root = Path(project_root) if project_root else path.resolve().parent.parent
    settings = Settings(raw=_load_yaml(path), source=path, project_root=root)
    _validate_settings(settings, require_runtime=require_runtime)
    return settings


def batch_type(profile: Dict[str, Any]) -> str:
    """batch 定義の起動方式。未指定は exe（従来の .exe 直接起動）。"""
    return str(profile.get("type") or "exe")


def batch_program_key(profile: Dict[str, Any]) -> str:
    """起動に必須の項目名。exe なら exe_path、java なら jar。"""
    return "jar" if batch_type(profile) == "java" else "exe_path"


def _validate_batch_profile(profile: Any, where: str, source: Path) -> None:
    """batch / batches.<名前> の型と項目名を検証する。"""
    if profile is None:
        return
    if not isinstance(profile, dict):
        raise ConfigError("%s: %s はマッピングで指定してください" % (source, where))
    unknown = sorted(set(profile) - BATCH_PROFILE_KEYS)
    if unknown:
        raise ConfigError("%s: %s に未知の項目があります: %s（綴り間違いの可能性。使える項目: %s）"
                          % (source, where, unknown, sorted(BATCH_PROFILE_KEYS)))
    if "type" in profile and profile["type"] not in BATCH_TYPES:
        raise ConfigError("%s: %s.type は %s のいずれかで指定してください（指定値: %r）"
                          % (source, where, " / ".join(BATCH_TYPES), profile["type"]))
    # 文字列を渡すと 1 文字ずつに分解されて起動コマンドへ並ぶため、リスト以外は止める
    for key in ("common_args", "jvm_args"):
        if key in profile and not isinstance(profile[key], list):
            raise ConfigError('%s: %s.%s はリストで指定してください（例: ["-Xmx1g"]）'
                              % (source, where, key))
    if "env_vars" in profile and profile["env_vars"] is not None \
            and not isinstance(profile["env_vars"], dict):
        raise ConfigError("%s: %s.env_vars はマッピングで指定してください" % (source, where))


def pg_host_port(db: Dict[str, Any]) -> Tuple[str, int]:
    """PostgreSQL の接続先を (ホスト, ポート) にする。server は host / host:port / [IPv6]:port。

    SQL Server 流の host,1433 は受け付けない。カンマを含んだままホスト名として
    名前解決に進み、原因の分かりにくい接続失敗になるため。
    """
    text = str(db.get("server") or "").strip()
    host, port_text = text, None
    if text.startswith("["):
        end = text.find("]")
        if end < 0 or (text[end + 1:] and not text[end + 1:].startswith(":")):
            raise ConfigError("database.server の書式が不正です: %s（IPv6 は [::1]:5432 の形）" % text)
        host = text[1:end]
        port_text = text[end + 2:] if text[end + 1:] else None
    elif text.count(":") == 1:
        host, port_text = text.split(":")
    if "," in host:
        raise ConfigError("PostgreSQL の database.server にカンマは書けません: %s"
                          "（ポートは host:5432 か port: 5432 で指定します）" % text)
    if not host.strip():
        raise ConfigError("database.server にホスト名がありません: %r" % text)
    port = db.get("port")
    if port_text is not None:
        if port is not None:
            raise ConfigError("ポートが database.server（%s）と database.port の両方に書かれています。"
                              "どちらか一方にしてください" % text)
        port = port_text
    if port is None:
        return host.strip(), PG_DEFAULT_PORT
    try:
        number = int(str(port).strip())
    except ValueError:
        number = 0
    if isinstance(port, bool) or not 0 < number < 65536:
        raise ConfigError("database のポート番号が不正です: %r" % port)
    return host.strip(), number


def _validate_database(db: Any, source: Path) -> None:
    """database の種類と項目名を検証する。"""
    if not isinstance(db, dict):
        return
    kind = db.get("type", "sqlserver")
    if kind not in DB_TYPES:
        raise ConfigError("%s: database.type は %s のいずれかで指定してください（指定値: %r）"
                          % (source, " / ".join(DB_TYPES), kind))
    allowed = DATABASE_POSTGRESQL_KEYS if kind == "postgresql" else DATABASE_SQLSERVER_KEYS
    unknown = sorted(set(db) - allowed)
    if unknown:
        lines = ["%s: database（type: %s）に使えない項目があります: %s" % (source, kind, unknown)]
        lines.extend("  - %s: %s" % (key, _DATABASE_KEY_HINTS[(kind, key)])
                     for key in unknown if (kind, key) in _DATABASE_KEY_HINTS)
        lines.append("  使える項目: %s" % sorted(allowed))
        raise ConfigError("\n".join(lines))
    for key in ("login_timeout_sec", "query_timeout_sec", "batch_start_margin_sec"):
        value = db.get(key)
        if key in db and (isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0):
            raise ConfigError("%s: database.%s は 0 以上の秒数で指定してください（指定値: %r）" % (source, key, value))
    if kind == "sqlserver":
        if str(db.get("auth", "sql")).lower() not in ("sql", "windows"):
            raise ConfigError("%s: database.auth は sql / windows のどちらかです（指定値: %r）"
                              % (source, db.get("auth")))
        return
    if db.get("sslmode", "prefer") not in PG_SSLMODES:
        raise ConfigError("%s: database.sslmode は %s のいずれかです（指定値: %r）"
                          % (source, " / ".join(PG_SSLMODES), db.get("sslmode")))
    for key in ("user", "password_env", "sslrootcert", "timezone", "search_path"):
        if key in db and not isinstance(db[key], str):
            raise ConfigError("%s: database.%s は文字列で指定してください（指定値: %r）" % (source, key, db[key]))
    if not db.get("user"):
        # 省略時に OS のユーザー名で繋ぐ libpq の流儀は、試験機ごとに接続ユーザーが変わるので採らない
        raise ConfigError("%s: PostgreSQL では database.user が必須です" % source)
    try:
        pg_host_port(db)
    except ConfigError as exc:
        raise ConfigError("%s: %s" % (source, exc))


def _validate_settings(s: Settings, require_runtime: bool = True) -> None:
    for key in ("env", "batch", "batches", "database", "paths", "log", "evidence", "excel", "folder_evidence"):
        if key == "batches" and s.raw.get(key) is None:
            continue  # 既存の batches:（空値）は未定義として扱う。
        if key in s.raw and not isinstance(s.raw[key], dict):
            raise ConfigError("%s: %s はマッピングで指定してください" % (s.sources[-1], key))
    _validate_database(s.raw.get("database"), s.sources[-1])
    _validate_batch_profile(s.raw.get("batch"), "batch", s.sources[-1])
    for name, profile in sorted(s.batches.items()):
        _validate_batch_profile(profile, "batches.%s" % name, s.sources[-1])
    missing: List[str] = []
    if require_runtime and not s.batch.get(batch_program_key(s.batch)):
        missing.append("batch.%s" % batch_program_key(s.batch))
    if not s.database.get("server"):
        missing.append("database.server")
    if not s.database.get("database"):
        missing.append("database.database")
    if require_runtime and not s.raw.get("paths"):
        missing.append("paths")
    if missing:
        raise ConfigError(f"{s.sources[-1]} に必須項目がありません: {', '.join(missing)}")

    if not require_runtime:
        return

    # folder_evidence.targets が paths に存在するか
    aliases = set(s.path_aliases)
    cfg = _folder_evidence_options(s.folder_evidence, "folder_evidence")
    unknown = [t for t in cfg.get("targets", []) if t not in aliases]
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
        cases_root: Optional[Path] = None,
    ) -> None:
        self.case_id = case_id
        self.name = name
        self.source = source
        self.cases_root = cases_root
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
# 下位の mapping も未知の項目は止める。mask を masks と書くと個人情報がそのまま
# Excel に出る、preview を preveiw と書くと内容の証跡が残らない —— どちらも
# 「書いたのに効かない」が黙って続く類の誤りで、AI が生成した定義で特に起きやすい。
CLEAN_DIR_KEYS = {"alias", "exclude"}
SQL_ITEM_KEYS = {"file", "inline"}
INPUT_FILE_KEYS = {"src", "dest_dir", "rename", "corrupt"}
CORRUPT_KEYS = {"mode", "ratio", "offset", "content"}
REPLACE_FILE_KEYS = {"src", "dest_dir", "name"}
SNAPSHOT_KEYS = {"tables"}
SNAPSHOT_TABLE_KEYS = {"name", "table", "sql", "mask"}
COLLECT_KEYS = {"files", "folder_evidence"}
COLLECT_FILE_KEYS = {"dir", "pattern", "preview", "recursive"}
# execute.expected_exit_code は以前ここで受け付けていたが、どこからも参照されず
# 何も判定していなかった（書いた人は終了コードを確認しているつもりになる = 偽 OK）。
# 終了コードの判定は assert.exit_code に一本化し、書かれていたら案内して止める。
EXECUTE_KEYS = {"batch", "args", "date"}
ASSERT_KEYS = {"exit_code", "job_status", "db", "files", "log"}
DB_ASSERT_KEYS = {"table", "expected", "key", "ignore_columns", "manual"}
FILE_ASSERT_KEYS = {"name", "actual", "expected", "encoding", "ignore_line_patterns",
                    "exists", "manual"}
FILE_ACTUAL_KEYS = {"dir", "pattern"}
FILE_EXISTS_KEYS = {"dir", "pattern", "count"}
LOG_ASSERT_KEYS = {"must_contain", "must_not_contain", "manual"}
EXIT_CODE_ASSERT_KEYS = {"expected", "manual"}
JOB_STATUS_ASSERT_KEYS = {"expected", "job", "manual"}
TEARDOWN_KEYS = {"sql", "clean_dirs"}
# Spring Batch の BatchStatus。綴り間違い（COMPLETE 等）は一致しようがなく、
# 実行後に「期待と違う」NG で初めて気づくことになるため、定義の段階で止める。
SPRING_BATCH_STATUSES = ("COMPLETED", "FAILED", "STOPPED", "ABANDONED",
                         "STARTING", "STARTED", "STOPPING", "UNKNOWN")


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

    job_status = assertions.get("job_status")
    if job_status is not None:
        expected_status = job_status
        if isinstance(job_status, dict):
            problems += _check_unknown_keys(
                job_status, JOB_STATUS_ASSERT_KEYS, "assert.job_status", path)
            problems += _check_manual(job_status.get("manual"), "assert.job_status.manual")
            expected_status = job_status.get("expected")
            job_name = job_status.get("job")
            if job_name is not None and (not isinstance(job_name, str) or not job_name.strip()):
                problems.append("assert.job_status.job は空でないジョブ名の文字列で指定してください")
        if not isinstance(expected_status, str) or expected_status not in SPRING_BATCH_STATUSES:
            problems.append(
                "assert.job_status は %s のいずれか、または {expected: ..., job: ジョブ名} で"
                "指定してください（指定値: %r）" % (" / ".join(SPRING_BATCH_STATUSES), expected_status))

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
            for key in ("must_contain", "must_not_contain"):
                words = log_spec.get(key)
                if words is not None and (not isinstance(words, list) or any(
                        not isinstance(w, str) or not w for w in words)):
                    problems.append("assert.log.%s は空でない文字列のリストで指定してください" % key)
            # 文言が 1 つも無い log 判定は「必須 0 件すべて検出」で必ず OK になる。
            # 何も確認していないのに合格行が Excel に残るため、目視（manual: true）以外は止める
            if log_spec and log_spec.get("manual") is not True \
                    and not log_spec.get("must_contain") and not log_spec.get("must_not_contain"):
                problems.append(
                    "assert.log に must_contain / must_not_contain がどちらもありません"
                    "（何も確認せずに OK になります。目視で確認するなら manual: true を指定してください）")

    return problems


def _validate_sql_items(items: Any, where: str, path: Path) -> List[str]:
    """setup.sql / sql_after_batches / teardown.sql の各項目を検証する。"""
    if items is None:
        return []
    if not isinstance(items, list):
        return ["%s はリストで指定してください" % where]
    problems: List[str] = []
    for i, item in enumerate(items):
        item_where = "%s[%d]" % (where, i)
        if isinstance(item, str):
            continue
        if not isinstance(item, dict):
            problems.append("%s は SQL 文字列か {file: ...} で指定してください: %r" % (item_where, item))
            continue
        problems += _check_unknown_keys(item, SQL_ITEM_KEYS, item_where, path)
        if not item.get("file") and not item.get("inline"):
            problems.append("%s に file / inline がありません" % item_where)
    return problems


def _validate_dir_specs(items: Any, where: str, path: Path) -> List[str]:
    """clean_dirs / remove_dirs の各項目（論理名 or {alias, exclude}）を検証する。"""
    if items is None:
        return []
    if not isinstance(items, list):
        return ["%s はリストで指定してください" % where]
    problems: List[str] = []
    for i, spec in enumerate(items):
        if isinstance(spec, dict):
            problems += _check_unknown_keys(spec, CLEAN_DIR_KEYS, "%s[%d]" % (where, i), path)
            if not spec.get("alias"):
                problems.append("%s[%d] に alias がありません: %r" % (where, i, spec))
        elif not isinstance(spec, str):
            problems.append("%s[%d] は論理名の文字列か {alias: ...} で指定してください: %r"
                            % (where, i, spec))
    return problems


def _check_strict_bool(spec: Dict[str, Any], key: str, where: str) -> List[str]:
    """preview 等の真偽値。"true" と書くと文字列になり、is True 判定で無視される。"""
    if key in spec and not isinstance(spec[key], bool):
        return ["%s.%s は true / false で指定してください（引用符は付けないでください。指定値: %r）"
                % (where, key, spec[key])]
    return []


def _validate_snapshot(snapshot: Any, path: Path) -> List[str]:
    if snapshot is None:
        return []
    if not isinstance(snapshot, dict):
        return ["snapshot はマッピングで指定してください"]
    problems = _check_unknown_keys(snapshot, SNAPSHOT_KEYS, "snapshot", path)
    tables = snapshot.get("tables")
    if tables is None:
        return problems
    if not isinstance(tables, list):
        return problems + ["snapshot.tables はリストで指定してください"]
    for i, spec in enumerate(tables):
        where = "snapshot.tables[%d]" % i
        if not isinstance(spec, dict):
            problems.append("%s はマッピングで指定してください: %r" % (where, spec))
            continue
        problems += _check_unknown_keys(spec, SNAPSHOT_TABLE_KEYS, where, path)
        if not spec.get("name") and not spec.get("table"):
            problems.append("%s に name がありません（Excel の表題と assert.db の対応付けに使います）" % where)
        if "sql" in spec and not isinstance(spec["sql"], str):
            problems.append("%s.sql は SQL 文字列で指定してください" % where)
        mask = spec.get("mask")
        if mask is not None and (not isinstance(mask, list) or any(not isinstance(m, str) for m in mask)):
            problems.append("%s.mask は列名のリストで指定してください" % where)
    return problems


def _validate_teardown(teardown: Any, path: Path) -> List[str]:
    if teardown is None:
        return []
    if not isinstance(teardown, dict):
        return ["teardown はマッピングで指定してください"]
    problems = _check_unknown_keys(teardown, TEARDOWN_KEYS, "teardown", path)
    problems += _validate_sql_items(teardown.get("sql"), "teardown.sql", path)
    problems += _validate_dir_specs(teardown.get("clean_dirs"), "teardown.clean_dirs", path)
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
        problems += _validate_sql_items(setup.get(key), "setup.%s" % key, path)

    problems += _validate_dir_specs(setup.get("clean_dirs"), "setup.clean_dirs", path)
    # remove_dirs はフォルダごと消すので exclude の意味が無く、orchestrator も
    # 論理名の文字列としてしか扱わない（辞書を渡すと実行時に落ちる）。文字列に限る
    remove_dirs = setup.get("remove_dirs")
    if remove_dirs is not None:
        if not isinstance(remove_dirs, list):
            problems.append("setup.remove_dirs はリストで指定してください")
        else:
            for i, spec in enumerate(remove_dirs):
                if not isinstance(spec, str) or not spec.strip():
                    problems.append("setup.remove_dirs[%d] は論理名の文字列で指定してください: %r"
                                    % (i, spec))

    for key, allowed in (("input_files", INPUT_FILE_KEYS), ("replace_files", REPLACE_FILE_KEYS)):
        specs = setup.get(key)
        if specs is not None and not isinstance(specs, list):
            problems.append("setup.%s はリストで指定してください" % key)
            continue
        for i, spec in enumerate(specs or []):
            where = "setup.%s[%d]" % (key, i)
            if not isinstance(spec, dict):
                problems.append("%s はマッピングで指定してください: %r" % (where, spec))
                continue
            problems += _check_unknown_keys(spec, allowed, where, path)
            if not spec.get("src"):
                problems.append("%s に src がありません" % where)
            else:
                for field in ("src", "rename", "name"):
                    problems += _check_relative(spec.get(field), "%s.%s" % (where, field))
            corrupt = spec.get("corrupt")
            if isinstance(corrupt, dict):
                problems += _check_unknown_keys(corrupt, CORRUPT_KEYS, where + ".corrupt", path)
            elif corrupt is not None and not isinstance(corrupt, str):
                problems.append("%s.corrupt はモード名か {mode: ...} で指定してください" % where)

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

    problems += _validate_snapshot(data.get("snapshot"), path)

    execute = data.get("execute") or {}
    if not isinstance(execute, dict):
        problems.append("execute はマッピングで指定してください")
        execute = {}
    if "expected_exit_code" in execute:
        problems.append(
            "execute.expected_exit_code は判定に使われません（何も確認しないまま合格に見えてしまうため"
            "廃止しました）。終了コードの判定は assert.exit_code に書いてください")
        execute = dict(execute)
        del execute["expected_exit_code"]
    problems += _check_unknown_keys(execute, EXECUTE_KEYS, "execute", path)
    args = execute.get("args")
    if args is not None and not isinstance(args, list):
        # 文字列を渡すと 1 文字ずつに分解されて .exe へ渡ってしまう
        problems.append('execute.args はリストで指定してください（例: ["--mode", "daily"]）。'
                        "文字列を渡すと 1 文字ずつ分解されます。指定値: %r" % args)

    collect = data.get("collect") or {}
    if isinstance(collect, dict):
        problems += _check_unknown_keys(collect, COLLECT_KEYS, "collect", path)
        if "folder_evidence" in collect:
            try:
                _folder_evidence_options(collect["folder_evidence"], "collect.folder_evidence")
            except ConfigError as exc:
                problems.append(str(exc))
        files = collect.get("files")
        if files is not None and not isinstance(files, list):
            problems.append("collect.files はリストで指定してください")
        for i, spec in enumerate(files if isinstance(files, list) else []):
            where = "collect.files[%d]" % i
            if not isinstance(spec, dict):
                problems.append("%s はマッピングで指定してください: %r" % (where, spec))
                continue
            problems += _check_unknown_keys(spec, COLLECT_FILE_KEYS, where, path)
            for key in ("preview", "recursive"):
                problems += _check_strict_bool(spec, key, where)
    else:
        problems.append("collect はマッピングで指定してください")

    problems += _validate_review(data.get("review"), path)
    problems += _validate_teardown(data.get("teardown"), path)

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
    all_yaml = sorted(p for p in list(cases_dir.rglob("*.yaml")) + list(cases_dir.rglob("*.yml"))
                      if p.name not in SETTINGS_NAMES + LOCAL_SETTINGS_NAMES)
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


def load_cases(cases_dir: Union[str, Path], only: Optional[List[str]] = None, tags: Optional[List[str]] = None,
               include_disabled: bool = False) -> List[TestCase]:
    """cases/ 配下の *.yaml を読み込む。only / tags で絞り込み可能。

    サブフォルダ名は暗黙のタグになる（cases/group_a/TC001.yaml なら「group_a」）。
    グループごとにフォルダを切っておけば --tag group_a でまとめて実行できる。

    include_disabled は validate / list 専用。生成直後のケースは enabled: false なので、
    これが無いと「作りかけを静的検査する手段」が無く、validate が OK を返しても
    実は何も検査していない状態になる。run では使わない（enabled は未レビューの印）。
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

    if include_disabled:
        return cases

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
        cases_root=cases_dir,
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
