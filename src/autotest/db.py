"""SQL Server アクセスと、取得値の Excel 向けフォーマット。

DB 値はユーザ指示により画像化せず、Excel のネイティブセルとして書き出す。
そのため「表示用に整形した文字列」を Table に詰めるのがこのモジュールの役割。

offline モード:
  pyodbc / SQL Server が無い環境（macOS での動作確認など）で
  fixtures/<case_id>/<phase>_<table>.csv を DB の代わりに読む。
  パイプライン全体（比較 → 判定 → Excel 出力）を実 DB 無しで検証できる。
"""

import csv
import datetime as dt
import decimal
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from .config import ConfigError, Settings
from .models import Table


class DbError(Exception):
    """DB 接続・クエリ実行の失敗。"""


# =============================================================================
# 値フォーマット
# =============================================================================


def format_value(value: Any, fmt: Dict[str, Any], column: str = "") -> Any:
    """DB の生値を Excel セルに入れる形へ整形する。

    数値は Excel 側でも数値のままにしたい場合があるため int はそのまま返す。
    Decimal は既定で文字列化する（Excel の float 丸めで桁落ちさせないため）。
    """
    if value is None:
        return fmt.get("null_text", "(NULL)")

    if isinstance(value, dt.datetime):
        return value.strftime(fmt.get("datetime", "%Y-%m-%d %H:%M:%S"))
    if isinstance(value, dt.date):
        return value.strftime(fmt.get("date", "%Y-%m-%d"))
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S")

    if isinstance(value, decimal.Decimal):
        places = (fmt.get("decimal_places") or {}).get(column)
        if places is not None:
            quant = decimal.Decimal(1).scaleb(-int(places))
            value = value.quantize(quant, rounding=decimal.ROUND_HALF_UP)
        return str(value) if fmt.get("decimal_as_text", True) else float(value)

    if isinstance(value, (bytes, bytearray, memoryview)):
        head = int(fmt.get("binary_head_bytes", 16))
        raw = bytes(value)
        hexed = "0x" + raw[:head].hex().upper()
        return hexed + ("..." if len(raw) > head else "")

    if isinstance(value, bool):
        return "1" if value else "0"

    if isinstance(value, str):
        return value.rstrip() if fmt.get("trim_trailing_space", True) else value

    return value


def apply_mask(row: Dict[str, Any], mask_columns: List[str]) -> Dict[str, Any]:
    """個人情報等をマスクする。証跡がそのまま共有されても問題ない状態にする。"""
    if not mask_columns:
        return row
    masked = dict(row)
    for col in mask_columns:
        if col in masked and masked[col] is not None:
            masked[col] = "***MASKED***"
    return masked


# =============================================================================
# クライアント
# =============================================================================


class DbClient(ABC):
    @abstractmethod
    def query(self, sql: str, params: Optional[List[Any]] = None) -> Tuple[List[str], List[List[Any]]]:
        """SELECT を実行し (列名, 行) を返す。"""

    @abstractmethod
    def execute_script(self, sql: str) -> None:
        """INSERT/DELETE 等を実行しコミットする。"""

    def close(self) -> None:  # pragma: no cover - 既定は何もしない
        pass

    # --- 共通処理 -----------------------------------------------------------
    # 判定に使うスナップショットの安全上限（メモリ保護）。これを超えたら
    # truncated_from に打ち切りを記録し、compare 側が「打ち切られたデータでは
    # 判定できない」と NG にする。Excel の表示行数（excel.max_db_rows）とは別物で、
    # 表示の絞り込みは excel.py 側の仕事。判定は常に全行で行う。
    JUDGE_ROW_CAP = 100000

    def snapshot(self, spec: Dict[str, Any], fmt: Dict[str, Any], title_prefix: str = "") -> Table:
        """スナップショット定義 1 件を Table に変換する。

        判定の元データになるため、表示都合の打ち切りはここでは行わない。
        """
        table_name = str(spec.get("name") or spec.get("table") or "UNKNOWN")
        sql = spec.get("sql") or f"SELECT * FROM {table_name}"
        columns, rows = self.query(sql)

        local_fmt = {**fmt, **(spec.get("format") or {})}
        mask_columns = list(spec.get("mask") or [])

        total = len(rows)
        max_rows = self.JUDGE_ROW_CAP
        if total > max_rows:
            rows = rows[:max_rows]

        formatted: List[List[Any]] = []
        for row in rows:
            as_dict = apply_mask(dict(zip(columns, row)), mask_columns)
            formatted.append([format_value(as_dict[c], local_fmt, c) for c in columns])

        return Table(
            title=f"{title_prefix}{table_name}",
            columns=columns,
            rows=formatted,
            truncated_from=total if total > max_rows else None,
            note=f"SQL: {sql.strip()}",
        )


class PyodbcClient(DbClient):
    """本番用。pyodbc 経由で SQL Server に接続する。"""

    def __init__(self, settings: Settings):
        try:
            import pyodbc  # noqa: PLC0415  遅延 import（offline モードでは不要なため）
        except ImportError as exc:  # pragma: no cover
            raise DbError(
                "pyodbc が import できません。`pip install pyodbc` と "
                "Microsoft ODBC Driver for SQL Server のインストールを確認してください。"
            ) from exc

        db = settings.database
        parts = [
            f"DRIVER={{{db.get('driver', 'ODBC Driver 18 for SQL Server')}}}",
            f"SERVER={db['server']}",
            f"DATABASE={db['database']}",
        ]
        if str(db.get("auth", "sql")).lower() == "windows":
            parts.append("Trusted_Connection=yes")
        else:
            parts.append(f"UID={db.get('user', '')}")
            parts.append(f"PWD={settings.db_password()}")
        parts.append("Encrypt=yes" if db.get("encrypt", True) else "Encrypt=no")
        if db.get("trust_server_certificate", True):
            parts.append("TrustServerCertificate=yes")

        conn_str = ";".join(parts)
        self._redacted = conn_str.replace(settings.db_password(), "****") if settings.db_password() else conn_str
        try:
            self.conn = pyodbc.connect(conn_str, timeout=int(db.get("login_timeout_sec", 15)))
        except Exception as exc:  # pragma: no cover - 環境依存
            raise DbError(f"SQL Server への接続に失敗しました: {exc}\n  接続文字列: {self._redacted}") from exc
        self.query_timeout = int(db.get("query_timeout_sec", 60))

    def query(self, sql: str, params: Optional[List[Any]] = None) -> Tuple[List[str], List[List[Any]]]:
        cur = self.conn.cursor()
        cur.timeout = self.query_timeout
        try:
            cur.execute(sql, params or [])
            if cur.description is None:
                return [], []
            columns = [d[0] for d in cur.description]
            rows = [list(r) for r in cur.fetchall()]
            return columns, rows
        except Exception as exc:
            raise DbError(f"クエリ実行に失敗しました: {exc}\n  SQL: {sql}") from exc
        finally:
            cur.close()

    def execute_script(self, sql: str) -> None:
        """GO 区切りに対応したスクリプト実行。"""
        cur = self.conn.cursor()
        cur.timeout = self.query_timeout
        try:
            for batch in _split_go_batches(sql):
                if batch.strip():
                    cur.execute(batch)
            self.conn.commit()
        except Exception as exc:
            self.conn.rollback()
            raise DbError(f"SQL スクリプトの実行に失敗しました: {exc}") from exc
        finally:
            cur.close()

    def close(self) -> None:
        try:
            self.conn.close()
        except Exception:  # pragma: no cover
            pass


class OfflineClient(DbClient):
    """DB 無しでパイプラインを検証するための代替。fixtures/ の CSV を返す。"""

    def __init__(self, fixtures_dir: Path, case_id: str):
        self.dir = Path(fixtures_dir) / case_id
        self.phase = "before"
        self._executed: List[str] = []

    def set_phase(self, phase: str) -> None:
        self.phase = phase

    def query(self, sql: str, params: Optional[List[Any]] = None) -> Tuple[List[str], List[List[Any]]]:
        table = _guess_table_name(sql)
        path = self.dir / f"{self.phase}_{table}.csv"
        if not path.exists():
            raise DbError(
                f"offline モードのフィクスチャがありません: {path}\n"
                f"  実 DB の代わりに読むファイルです。ヘッダ付き CSV を配置してください。"
            )
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            rows = list(reader)
        if not rows:
            return [], []
        return rows[0], [list(r) for r in rows[1:]]

    def execute_script(self, sql: str) -> None:
        self._executed.append(sql)  # offline では記録のみ


def _split_go_batches(sql: str) -> List[str]:
    """SSMS 形式の GO 区切りでスクリプトを分割する。"""
    batches: List[str] = []
    current: List[str] = []
    for line in sql.splitlines():
        if line.strip().upper() == "GO":
            batches.append("\n".join(current))
            current = []
        else:
            current.append(line)
    batches.append("\n".join(current))
    return batches


def _guess_table_name(sql: str) -> str:
    """offline モードのフィクスチャ名決定用に SQL から FROM 句のテーブル名を拾う。"""
    tokens = sql.replace("\n", " ").split()
    for i, tok in enumerate(tokens):
        if tok.upper() == "FROM" and i + 1 < len(tokens):
            return tokens[i + 1].strip("[]\"'").split(".")[-1]
    return "UNKNOWN"


def list_installed_drivers() -> List[str]:
    """この端末にインストール済みの ODBC ドライバ名を返す。

    「settings.yaml の driver 名が実際に入っているものと違う」は
    接続失敗の最頻原因なので、失敗時に候補を提示できるようにする。
    """
    try:
        import pyodbc  # noqa: PLC0415
    except ImportError:
        return []
    try:
        return list(pyodbc.drivers())
    except Exception:
        return []


def build_connection_string(settings: Settings, timeout_sec: Optional[int] = None) -> Tuple[str, str]:
    """(接続文字列, パスワードを伏せた表示用文字列) を返す。"""
    db = settings.database
    parts = [
        "DRIVER={%s}" % db.get("driver", "ODBC Driver 18 for SQL Server"),
        "SERVER=%s" % db.get("server", ""),
        "DATABASE=%s" % db.get("database", ""),
    ]
    password = ""
    if str(db.get("auth", "sql")).lower() == "windows":
        parts.append("Trusted_Connection=yes")
    else:
        password = settings.db_password()
        parts.append("UID=%s" % db.get("user", ""))
        parts.append("PWD=%s" % password)
    parts.append("Encrypt=yes" if db.get("encrypt", True) else "Encrypt=no")
    if db.get("trust_server_certificate", True):
        parts.append("TrustServerCertificate=yes")

    conn_str = ";".join(parts)
    shown = conn_str.replace("PWD=%s" % password, "PWD=********") if password else conn_str
    return conn_str, shown


def diagnose_connection_error(exc: Exception) -> List[str]:
    """接続例外から、次に何を確認すべきかの助言を組み立てる。"""
    text = str(exc)
    hints: List[str] = []

    if "IM002" in text or "Data source name not found" in text:
        hints.append("settings.yaml の database.driver が、この端末に入っている名前と一致していません。")
        drivers = list_installed_drivers()
        if drivers:
            hints.append("インストール済みのドライバ: %s" % drivers)
            hints.append("この中の SQL Server 用のものを database.driver にそのまま書いてください。")
        else:
            hints.append("ODBC ドライバが 1 つも見つかりません。"
                         "Microsoft ODBC Driver for SQL Server を導入してください。")
    elif "08001" in text or "Named Pipes" in text or "TCP Provider" in text:
        hints.append("サーバへ到達できていません。次を確認してください:")
        hints.append("  - サーバ名 / IP とポート番号（例 SQLSRV01,1433）")
        hints.append("  - セキュリティグループ・ファイアウォールで 1433 が開いているか")
        hints.append("  - SQL Server 側で TCP/IP プロトコルが有効か")
    elif "28000" in text or "Login failed" in text:
        hints.append("サーバには到達できていますが、認証で拒否されました:")
        hints.append("  - ユーザー名 / パスワード（環境変数 %s）"
                     % "AUTOTEST_DB_PASSWORD")
        hints.append("  - Windows 認証なら database.auth を windows にする")
    elif "timeout" in text.lower() or "HYT00" in text:
        hints.append("接続がタイムアウトしました。到達できないホストか、"
                     "ファイアウォールがパケットを破棄している可能性があります。")
    elif "SSL" in text or "certificate" in text.lower():
        hints.append("TLS 証明書の検証で失敗しています。"
                     "検証環境なら database.trust_server_certificate: true を試してください。")
    return hints


class NullClient(DbClient):
    """--dry-run 用。DB へ一切接続しない。

    dry-run は「何も触らずに流れだけ確認する」ためのモードなので、
    実 DB への接続も行わない。接続を試みると、到達できないサーバ相手に
    TCP ハンドシェイクで長時間ブロックし、ハングしたように見える
    （login_timeout はハンドシェイク前の段階を必ずしもカバーしない）。
    """

    def query(self, sql: str, params: Optional[List[Any]] = None) -> Tuple[List[str], List[List[Any]]]:
        return [], []

    def execute_script(self, sql: str) -> None:
        pass


def create_client(settings: Settings, offline: bool, case_id: str, dry_run: bool = False) -> DbClient:
    if dry_run:
        return NullClient()
    if offline:
        return OfflineClient(settings.project_root / "fixtures", case_id)
    return PyodbcClient(settings)


def read_expected_csv(path: Path) -> Table:
    """期待値 CSV を Table として読み込む。"""
    if not path.exists():
        raise ConfigError(f"期待値ファイルがありません: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return Table(title=path.stem, columns=[], rows=[])
    return Table(title=path.stem, columns=rows[0], rows=[list(r) for r in rows[1:]])
