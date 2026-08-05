"""SQL Server アクセスと、取得値の文字列化。

DB 値は画像化せず Excel のネイティブセルとして書き出す。値は整形せず、
情報を落とさずに文字列化するだけ（DB にあるものをそのまま出す）。
丸めや桁の省略は「表示のための加工」であり、判定より前に適用すると
値が違うのに一致とみなされるため行わない。

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


NULL_TEXT = "(NULL)"
MASKED_TEXT = "***MASKED***"


def to_text(value: Any) -> str:
    """DB の値を「情報を落とさずに」文字列化する。

    ここでの整形は一切行わない。丸め・切り捨て・桁の省略をすると、
    表示は読みやすくなる代わりに、値が違うのに同じに見える状態が生まれ、
    比較で偽 OK になる（実際に mask / binary 16 バイト / 秒精度 /
    Decimal 丸めの 4 パターンで発生していた）。

    「DB にあるものを、そのまま CSV に出す」が原則。
    """
    if value is None:
        return NULL_TEXT

    if isinstance(value, dt.datetime):
        # マイクロ秒まで残す。秒で切ると 09:00:00.111 と 09:00:00.999 が同一になる
        return value.isoformat(sep=" ")
    if isinstance(value, (dt.date, dt.time)):
        return value.isoformat()

    if isinstance(value, decimal.Decimal):
        # 丸めない。表示桁で丸めると 100.004 と 100.001 が同一になる
        return str(value)

    if isinstance(value, (bytes, bytearray, memoryview)):
        # 全バイトを 16 進で出す。先頭 N バイトで打ち切ると
        # 後半だけ異なる BLOB を同一とみなしてしまう
        return "0x" + bytes(value).hex().upper()

    if isinstance(value, bool):
        return "1" if value else "0"

    # 文字列は前後の空白も含めてそのまま。固定長・EDI では空白に意味がある
    return str(value)


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

    def snapshot(self, spec: Dict[str, Any], title_prefix: str = "") -> Table:
        """スナップショット定義 1 件を Table に変換する。

        値は情報を落とさずに文字列化するだけで、整形もマスクも行わない。
        いずれも「表示のための加工」であり、判定より手前で適用すると
        値が違うのに一致とみなされる（偽 OK）。マスクは Excel へ書く
        直前に適用する（mask_columns として Table に持たせる）。
        """
        table_name = str(spec.get("name") or spec.get("table") or "UNKNOWN")
        sql = spec.get("sql") or f"SELECT * FROM {table_name}"
        columns, rows = self.query(sql)

        total = len(rows)
        max_rows = self.JUDGE_ROW_CAP
        if total > max_rows:
            rows = rows[:max_rows]

        return Table(
            title=f"{title_prefix}{table_name}",
            columns=columns,
            rows=[[to_text(v) for v in row] for row in rows],
            truncated_from=total if total > max_rows else None,
            note=f"SQL: {sql.strip()}",
            mask_columns=list(spec.get("mask") or []),
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
        # クエリのタイムアウトは Connection の属性。Cursor には timeout が無く、
        # cur.timeout = ... とすると AttributeError になる（pyodbc の仕様）。
        # 0 は「無制限」の意味になるため、未設定時も既定値を入れておく。
        self.query_timeout = int(db.get("query_timeout_sec", 60))
        try:
            self.conn.timeout = self.query_timeout
        except Exception:  # pragma: no cover - 古い pyodbc では未対応の場合がある
            pass

    def query(self, sql: str, params: Optional[List[Any]] = None) -> Tuple[List[str], List[List[Any]]]:
        cur = self.conn.cursor()
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


class LockHolder:
    """batch の実行中だけ別接続でロックを保持する。

    「DB 更新に失敗したとき batch がどう振る舞うか」は手作業では再現しづらい。
    SSMS で手動トランザクションを張ったまま batch を起動する、という作業を
    自動化するためのもの。ロック待ちタイムアウト・デッドロック・更新失敗時の
    ロールバックを、決まった手順で毎回同じように起こせる。

    with 文を抜けるとき必ずロールバックする。保持したまま終わると
    後続ケースや他の作業が全部止まるため、解放は例外時も必ず行う。
    """

    def __init__(self, settings: Settings, sql: str):
        self.settings = settings
        self.sql = sql
        self.conn = None

    def __enter__(self) -> "LockHolder":
        import pyodbc  # noqa: PLC0415

        conn_str, _shown = build_connection_string(self.settings)
        self.conn = pyodbc.connect(conn_str, timeout=int(
            self.settings.database.get("login_timeout_sec", 15)))
        cur = self.conn.cursor()
        try:
            # BEGIN TRAN したまま commit しないことでロックを保持する
            cur.execute(self.sql)
        finally:
            cur.close()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.conn is None:
            return
        try:
            self.conn.rollback()
        except Exception:
            pass
        try:
            self.conn.close()
        except Exception:
            pass
        self.conn = None


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


def server_now(client: "DbClient") -> Optional[dt.datetime]:
    """DB サーバ側の現在時刻を取得する。取れなければ None。

    テスト機の時計ではなく DB の時計を使う。両者にずれがあると、
    batch が更新した行を取りこぼす（＝異常の見逃し）ため。
    SYSDATETIME() が使えない環境向けに GETDATE() へフォールバックする。
    """
    for sql in ("SELECT SYSDATETIME()", "SELECT GETDATE()"):
        try:
            _columns, rows = client.query(sql)
        except Exception:
            continue
        if not rows or not rows[0]:
            continue
        value = rows[0][0]
        if isinstance(value, dt.datetime):
            return value
        for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
            try:
                return dt.datetime.strptime(str(value)[:26 if "%f" in fmt else 19], fmt)
            except ValueError:
                continue
    return None


def format_sql_datetime(value: dt.datetime) -> str:
    """SQL Server のリテラルとして埋め込める形式にする（ミリ秒まで）。

    ISO 8601 の T 区切り（yyyy-mm-ddThh:mi:ss.mmm）を使う。
    空白区切りの 'yyyy-mm-dd hh:mi:ss' は datetime 型では
    SET LANGUAGE / DATEFORMAT の影響を受け、環境によっては
    「文字列から日付への変換に失敗（SQLSTATE 22007 / エラー 241）」になる。
    T 区切りは言語設定に依存しない。
    """
    return value.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]


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


def parse_server(server: str) -> Tuple[str, Optional[int], Optional[str]]:
    """SQL Server の接続先文字列を (ホスト, ポート, 名前付きインスタンス) に分解する。

    SQL Server の記法はポート区切りがコロンではなくカンマ:
        SQLSRV01              -> 既定インスタンス。既定ポート 1433
        SQLSRV01,1433         -> ポート明示
        SQLSRV01\\SQLEXPRESS  -> 名前付きインスタンス。ポートは動的で、
                                 SQL Server Browser (UDP 1434) 経由で解決される
    """
    text = str(server).strip()
    instance = None
    port = None

    if "," in text:
        text, port_text = text.split(",", 1)
        try:
            port = int(port_text.strip())
        except ValueError:
            port = None
    if "\\" in text:
        text, instance = text.split("\\", 1)

    host = text.strip()
    if port is None and instance is None:
        port = 1433  # 既定インスタンスの既定ポート
    return host, port, instance


def check_tcp_reachable(host: str, port: int, timeout_sec: float = 5.0) -> Tuple[bool, str]:
    """TCP レベルで到達できるかを確認する。

    ODBC のエラーメッセージは原因が読み取りにくいため、先に素の TCP で
    「そもそもネットワークが通っているか」を切り分ける。ここが NG なら
    ドライバや認証情報をいくら見直しても無駄だと分かる。
    """
    import socket  # noqa: PLC0415

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout_sec)
    try:
        sock.connect((host, port))
        return True, ""
    except socket.timeout:
        return False, "タイムアウト（%.0f 秒）: パケットが破棄されている可能性（ファイアウォール/セキュリティグループ）" % timeout_sec
    except socket.gaierror as exc:
        return False, "ホスト名を解決できません: %s" % exc
    except ConnectionRefusedError:
        return False, "接続を拒否されました: ホストには届いているが、そのポートで待ち受けていない"
    except OSError as exc:
        return False, str(exc)
    finally:
        sock.close()


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


def diagnose_password_env(env_name: str, value: Optional[str]) -> List[str]:
    """パスワード環境変数の状態を診断する。値そのものは決して返さない。

    setx は「新しいプロセスにしか効かない」ため設定したつもりで未反映、
    cmd では % & ^ 等が解釈されて値が壊れる、という事故が起きやすい。
    """
    if value is None:
        return [
            "[NG] このプロセスからは見えていません。",
            "     setx で設定した場合、既存のコンソールには反映されません。",
            "     PowerShell を開き直すか、次で直接確認してください:",
            '       [Environment]::GetEnvironmentVariable("%s", "User")' % env_name,
        ]

    lines = ["[OK] 設定済み（%d 文字）" % len(value)]
    if value == "":
        return ["[NG] 空文字が設定されています。値の指定を確認してください。"]

    # cmd の setx はこれらを解釈するため、意図した値と違って保存されることがある
    risky = [c for c in "%&^<>|" if c in value]
    if risky:
        lines.append("[警告] 特殊文字 %s を含みます。cmd の setx で設定した場合、" % "".join(risky))
        lines.append("       値が壊れている可能性があります。PowerShell から設定し直せます:")
        lines.append('         [Environment]::SetEnvironmentVariable("%s", \'<パスワード>\', "User")' % env_name)
    if value != value.strip():
        lines.append("[警告] 前後に空白が含まれています。引用符の付け方を確認してください。")
    if value.startswith('"') and value.endswith('"'):
        lines.append("[警告] 値が引用符で囲まれたまま保存されています。引用符も値の一部になっています。")
    return lines


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
    elif "22007" in text or "(241)" in text:
        hints.append("日付時刻の文字列を DB 側で変換できませんでした。")
        hints.append("  - 比較対象の列が datetime 型か確認してください。")
        hints.append("    char/varchar に 'yyyyMMdd' 等で入っている場合は、書式を合わせます:")
        hints.append("      WHERE UPDATE_YMD >= '{batch_start:%Y%m%d}'")
        hints.append("  - 実際に送られた SQL は run.log と証跡 Excel に記録されています。")
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
