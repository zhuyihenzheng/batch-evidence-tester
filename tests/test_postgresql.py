# -*- coding: utf-8 -*-
"""PostgreSQL 対応の回帰テスト（内蔵クライアント pgwire / PostgresClient / 設定 / dbcheck / gen の書き方）。

実サーバが無くても動くよう、プロトコルは偽サーバで確かめる。実サーバでの確認は
test_postgresql_live.py（環境変数 AUTOTEST_PG_TEST_URL があるときだけ動く）。

守っていること:
  - NULL（長さ -1）と空文字（長さ 0）を取り違えない
  - SSL を要求したのに断られたら平文に落とさない
  - 未対応の認証方式・想定外の応答は推測で続けずに止める
  - gen は N'...' を使わない（PostgreSQL では末尾の空白が落ちる）

  python -m unittest tests.test_postgresql -v
"""

import base64
import datetime as dt
import hashlib
import hmac
import io
import os
import shutil
import socket
import struct
import sys
import tempfile
import threading
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import cli, pgwire  # noqa: E402
from autotest import db as db_mod  # noqa: E402
from autotest.config import ConfigError, Settings, load_settings, pg_host_port  # noqa: E402


# =============================================================================
# 偽サーバ
# =============================================================================


def _cstr(text):
    return text.encode("utf-8") + b"\0"


def _msg(kind, payload=b""):
    return kind + struct.pack("!i", len(payload) + 4) + payload


def auth(code, extra=b""):
    return _msg(b"R", struct.pack("!i", code) + extra)


def ready(status=b"I"):
    return _msg(b"Z", status)


def row_desc(*names):
    return _msg(b"T", struct.pack("!h", len(names)) + b"".join(
        _cstr(n) + struct.pack("!IhIhih", 0, 0, 25, -1, -1, 0) for n in names))


def data_row(*values):
    parts = [struct.pack("!h", len(values))]
    for value in values:
        if value is None:
            parts.append(struct.pack("!i", -1))
        else:
            raw = value.encode("utf-8")
            parts.append(struct.pack("!i", len(raw)) + raw)
    return _msg(b"D", b"".join(parts))


def complete(tag):
    return _msg(b"C", _cstr(tag))


def fields_msg(kind, **fields):
    return _msg(kind, b"".join(k.encode("ascii") + _cstr(v) for k, v in fields.items()) + b"\0")


def recv_exact(conn, size):
    buf = b""
    while len(buf) < size:
        chunk = conn.recv(size - len(buf))
        if not chunk:
            raise EOFError("client closed")
        buf += chunk
    return buf


def read_msg(conn):
    kind = recv_exact(conn, 1)
    length = struct.unpack("!i", recv_exact(conn, 4))[0]
    return kind, recv_exact(conn, length - 4)


def read_startup(conn, ssl_answer=b"N"):
    """起動メッセージを読む。SSLRequest が来たら ssl_answer で答えて読み直す。"""
    length = struct.unpack("!i", recv_exact(conn, 4))[0]
    body = recv_exact(conn, length - 4)
    if struct.unpack("!i", body[:4])[0] == pgwire.SSL_REQUEST_CODE:
        conn.sendall(ssl_answer)
        return read_startup(conn, ssl_answer)
    parts = body[4:].split(b"\0")
    return {parts[i].decode("utf-8"): parts[i + 1].decode("utf-8") for i in range(0, len(parts) - 2, 2)}


class FakeServer(object):
    """1 接続だけ受ける偽の PostgreSQL。handler(server, conn) がやり取りを書く。"""

    def __init__(self, handler):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(1)
        self.port = self.sock.getsockname()[1]
        self.handler = handler
        self.startup = None
        self.received = []
        self.error = None
        self.thread = threading.Thread(target=self._run)
        self.thread.daemon = True
        self.thread.start()

    def _run(self):
        try:
            conn, _addr = self.sock.accept()
            conn.settimeout(5)
            try:
                self.handler(self, conn)
            finally:
                conn.close()
        except Exception as exc:  # 偽サーバ側の失敗はテスト側で必要に応じて確認する
            self.error = exc
        finally:
            self.sock.close()

    def connect(self, **kwargs):
        kwargs.setdefault("sslmode", "disable")
        return pgwire.Connection("127.0.0.1", self.port, kwargs.pop("user", "app"),
                                 kwargs.pop("password", "pw"), kwargs.pop("database", "BATCH_DB"),
                                 connect_timeout=5, io_timeout=5, **kwargs)


def trust_then(*responses):
    """trust 認証のあと、Query を 1 つ受けるごとに responses を 1 つずつ返す偽サーバの処理。"""
    def handler(server, conn):
        server.startup = read_startup(conn)
        conn.sendall(auth(0) + _msg(b"S", _cstr("server_version") + _cstr("16.2"))
                     + _msg(b"K", struct.pack("!ii", 4242, 1)) + ready())
        for response in responses:
            kind, body = read_msg(conn)
            server.received.append((kind, body))
            conn.sendall(response)
        try:
            server.received.append(read_msg(conn))   # Terminate
        except (EOFError, OSError):
            pass
    return handler


# =============================================================================
# プロトコル
# =============================================================================


class TestStartupAndAuth(unittest.TestCase):
    def test_fixed_session_settings_are_sent(self):
        """値の文字表現をサーバの既定値に左右させない設定を、毎回起動時に送ること。"""
        server = FakeServer(trust_then())
        conn = server.connect(params=[("statement_timeout", "60000"), ("TimeZone", "Asia/Tokyo")])
        conn.close()
        server.thread.join(5)
        sent = server.startup
        self.assertEqual((sent["user"], sent["database"]), ("app", "BATCH_DB"))
        for key, value in (("client_encoding", "UTF8"), ("DateStyle", "ISO"), ("bytea_output", "hex"),
                           ("extra_float_digits", "3"), ("statement_timeout", "60000"), ("TimeZone", "Asia/Tokyo")):
            self.assertEqual(sent.get(key), value, key)
        self.assertEqual(conn.auth_method, "trust")
        self.assertEqual(conn.server_params.get("server_version"), "16.2")

    def test_cleartext_password(self):
        def handler(server, conn):
            read_startup(conn)
            conn.sendall(auth(3))
            server.received.append(read_msg(conn))
            conn.sendall(auth(0) + ready())
        server = FakeServer(handler)
        conn = server.connect(password="secret")
        conn.close()
        server.thread.join(5)
        self.assertEqual(server.received[0], (b"p", b"secret\0"))
        self.assertEqual(conn.auth_method, "password")

    def test_md5_password(self):
        salt = b"\x01\x02\x03\x04"

        def handler(server, conn):
            read_startup(conn)
            conn.sendall(auth(5, salt))
            server.received.append(read_msg(conn))
            conn.sendall(auth(0) + ready())
        server = FakeServer(handler)
        server.connect(user="batch", password="md5pw").close()
        server.thread.join(5)
        inner = hashlib.md5(b"md5pwbatch").hexdigest()
        expected = "md5" + hashlib.md5(inner.encode("ascii") + salt).hexdigest()
        self.assertEqual(server.received[0], (b"p", _cstr(expected)))

    def scram_handler(self, password, tamper=False, skip_final=False):
        salt, iterations = b"salt-1234", 4096

        def handler(server, conn):
            read_startup(conn)
            conn.sendall(auth(10, _cstr("SCRAM-SHA-256") + b"\0"))
            _kind, body = read_msg(conn)
            end = body.index(b"\0")
            size = struct.unpack("!i", body[end + 1:end + 5])[0]
            first = body[end + 5:end + 5 + size].decode("utf-8")
            first_bare = first[3:]
            nonce = dict(p.split("=", 1) for p in first_bare.split(","))["r"] + "srv"
            server_first = "r=%s,s=%s,i=%d" % (nonce, base64.b64encode(salt).decode("ascii"), iterations)
            conn.sendall(auth(11, server_first.encode("utf-8")))
            _kind, body = read_msg(conn)
            without_proof, proof = body.decode("utf-8").rsplit(",p=", 1)
            salted = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
            client_key = hmac.new(salted, b"Client Key", hashlib.sha256).digest()
            stored_key = hashlib.sha256(client_key).digest()
            message = ",".join((first_bare, server_first, without_proof)).encode("utf-8")
            signature = hmac.new(stored_key, message, hashlib.sha256).digest()
            recovered = bytes(a ^ b for a, b in zip(base64.b64decode(proof), signature))
            if hashlib.sha256(recovered).digest() != stored_key:
                conn.sendall(fields_msg(b"E", S="FATAL", C="28P01", M="password authentication failed"))
                return
            if skip_final:
                conn.sendall(auth(0) + ready())
                return
            server_key = hmac.new(salted, b"Server Key", hashlib.sha256).digest()
            verifier = b"\0" * 32 if tamper else hmac.new(server_key, message, hashlib.sha256).digest()
            conn.sendall(auth(12, b"v=" + base64.b64encode(verifier)) + auth(0) + ready())
        return handler

    def test_scram_sha_256_round_trip(self):
        server = FakeServer(self.scram_handler("scram-pw"))
        conn = server.connect(password="scram-pw")
        self.assertEqual(conn.auth_method, "SCRAM-SHA-256")
        conn.close()

    def test_scram_wrong_password_is_rejected_by_server(self):
        server = FakeServer(self.scram_handler("right"))
        with self.assertRaises(pgwire.PgError) as ctx:
            server.connect(password="wrong")
        self.assertEqual(ctx.exception.sqlstate, "28P01")

    def test_scram_server_signature_is_verified(self):
        """署名の合わないサーバ（本物でない可能性）とは接続を確立しないこと。"""
        server = FakeServer(self.scram_handler("pw", tamper=True))
        with self.assertRaises(pgwire.PgError) as ctx:
            server.connect(password="pw")
        self.assertIn("署名", str(ctx.exception))

    def test_auth_ok_without_scram_final_is_rejected(self):
        server = FakeServer(self.scram_handler("pw", skip_final=True))
        with self.assertRaises(pgwire.PgError) as ctx:
            server.connect(password="pw")
        self.assertIn("SCRAM 認証が完了する前", str(ctx.exception))

    def test_scram_matches_rfc7677_vector(self):
        scram = pgwire._Scram("pencil", nonce="rOprNGfwEbeRWgbNEkqO", user="user")
        self.assertEqual(scram.client_first(), b"n,,n=user,r=rOprNGfwEbeRWgbNEkqO")
        final = scram.client_final(b"r=rOprNGfwEbeRWgbNEkqO%hvYDpWUa2RaTCAfuxFIlj)hNlF$k0,"
                                   b"s=W22ZaJ0SNY7soEsUEjb6gQ==,i=4096")
        self.assertEqual(final, b"c=biws,r=rOprNGfwEbeRWgbNEkqO%hvYDpWUa2RaTCAfuxFIlj)hNlF$k0,"
                                b"p=dHzbZapWIk4jUhN+Ute9ytag9zjfMHgsqmmiz7AndVQ=")
        scram.verify_server(b"v=6rriTRBi23WpRR/wtup+mMhUZUn/dB5nLTJRsjl95G4=")
        self.assertTrue(scram.verified)

    def test_saslprep_matches_server_normalization(self):
        self.assertEqual(pgwire.saslprep("plain-ascii"), "plain-ascii")
        self.assertEqual(pgwire.saslprep("ｐａｓｓ東京"), "pass東京", "全角英字は NFKC で半角になる")
        self.assertEqual(pgwire.saslprep("a\u00a0b"), "a b", "非 ASCII の空白は空白に")
        self.assertEqual(pgwire.saslprep("a\u00adb"), "ab", "ソフトハイフンは消える")
        self.assertEqual(pgwire.saslprep("\u3000\u0007x"), "\u3000\u0007x", "禁止文字を含むなら元のまま")

    def test_unsupported_auth_method_names_pg_hba(self):
        def handler(server, conn):
            read_startup(conn)
            conn.sendall(auth(7))
        server = FakeServer(handler)
        with self.assertRaises(pgwire.PgError) as ctx:
            server.connect()
        self.assertIn("GSSAPI", str(ctx.exception))
        self.assertIn("pg_hba.conf", str(ctx.exception))

    def test_password_requested_but_not_configured(self):
        def handler(server, conn):
            read_startup(conn)
            conn.sendall(auth(3))
        server = FakeServer(handler)
        with self.assertRaises(pgwire.PgError) as ctx:
            server.connect(password=None)
        self.assertIn("password_env", str(ctx.exception))

    def test_ssl_refused_under_require_never_falls_back_to_plaintext(self):
        def handler(server, conn):
            server.startup = read_startup(conn, ssl_answer=b"N")
        server = FakeServer(handler)
        with self.assertRaises(pgwire.PgError) as ctx:
            server.connect(sslmode="require")
        self.assertIn("SSL を受け付けません", str(ctx.exception))
        server.thread.join(5)
        self.assertIsNone(server.startup, "平文で起動メッセージ（ユーザー名）を送っている")

    def test_ssl_refused_under_prefer_continues_in_plaintext(self):
        server = FakeServer(trust_then())
        conn = server.connect(sslmode="prefer")
        self.assertEqual(conn.ssl_version, "")
        conn.close()

    def test_unknown_sslmode_is_rejected(self):
        with self.assertRaises(pgwire.PgError):
            pgwire.Connection("127.0.0.1", 1, "u", None, "d", sslmode="allow")


class TestQuery(unittest.TestCase):
    def test_null_and_empty_string_are_distinct(self):
        server = FakeServer(trust_then(row_desc("a", "b", "c", "d") + data_row(None, "", "(NULL)", "ab  ")
                                       + complete("SELECT 1") + ready()))
        conn = server.connect()
        result = conn.execute("SELECT …")[0]
        self.assertEqual(result.columns, ["a", "b", "c", "d"])
        self.assertEqual(result.rows, [[None, "", "(NULL)", "ab  "]])
        conn.close()

    def test_error_mid_script_is_raised_after_ready_and_connection_stays_usable(self):
        server = FakeServer(trust_then(
            complete("INSERT 0 1")
            + fields_msg(b"E", S="ERROR", C="42P01", M='relation "t" does not exist', D="詳細", H="ヒント", P="15")
            + ready(b"I"),
            row_desc("x") + data_row("1") + complete("SELECT 1") + ready()))
        conn = server.connect()
        with self.assertRaises(pgwire.PgError) as ctx:
            conn.execute("INSERT …; SELECT * FROM t")
        error = ctx.exception
        self.assertEqual(error.sqlstate, "42P01")
        for part in ('relation "t" does not exist', "詳細", "ヒント", "15 文字目"):
            self.assertIn(part, str(error))
        self.assertEqual(conn.execute("SELECT 1")[0].rows, [["1"]], "エラーの後も同じ接続で続けられること")
        conn.close()

    def test_transaction_status_byte_is_tracked(self):
        server = FakeServer(trust_then(complete("BEGIN") + ready(b"T"),
                                       fields_msg(b"E", S="ERROR", C="22012", M="division by zero") + ready(b"E")))
        conn = server.connect()
        conn.execute("BEGIN")
        self.assertEqual(conn.transaction_status, "T")
        with self.assertRaises(pgwire.PgError):
            conn.execute("SELECT 1/0")
        self.assertEqual(conn.transaction_status, "E")
        conn.close()

    def test_results_per_statement_and_notices(self):
        server = FakeServer(trust_then(
            complete("DELETE 2") + fields_msg(b"N", S="NOTICE", C="00000", M="hello")
            + row_desc("n") + data_row("7") + complete("SELECT 1") + ready()))
        conn = server.connect()
        results = conn.execute("DELETE …; SELECT 7")
        self.assertEqual([(r.columns, r.command) for r in results], [(None, "DELETE 2"), (["n"], "SELECT 1")])
        self.assertIn("hello", conn.notices[0])
        conn.close()

    def test_connection_lost_mid_response(self):
        def handler(server, conn):
            read_startup(conn)
            conn.sendall(auth(0) + ready())
            read_msg(conn)
            conn.sendall(row_desc("a") + data_row("1")[:5])    # 途中で切る
        server = FakeServer(handler)
        conn = server.connect()
        with self.assertRaises(pgwire.PgError) as ctx:
            conn.execute("SELECT 1")
        self.assertIn("接続が切れました", str(ctx.exception))
        with self.assertRaises(pgwire.PgError):
            conn.execute("SELECT 1")
        conn.close()

    def test_unexpected_message_breaks_the_connection(self):
        server = FakeServer(trust_then(_msg(b"?", b"junk") + ready()))
        conn = server.connect()
        with self.assertRaises(pgwire.PgError) as ctx:
            conn.execute("SELECT 1")
        self.assertIn("プロトコルエラー", str(ctx.exception))
        with self.assertRaises(pgwire.PgError):
            conn.execute("SELECT 1")
        conn.close()

    def test_nul_in_sql_is_rejected_before_sending(self):
        server = FakeServer(trust_then())
        conn = server.connect()
        with self.assertRaises(pgwire.PgError):
            conn.execute("SELECT '\0'")
        conn.close()


# =============================================================================
# PostgresClient / LockHolder / 時刻
# =============================================================================


def pg_settings(**database):
    db = {"type": "postgresql", "server": "pg01", "database": "BATCH_DB", "user": "app"}
    db.update(database)
    return Settings(raw={"batch": {"exe_path": "x"}, "paths": {"log_dir": "./log"}, "database": db},
                    source=Path("settings.yaml"), project_root=Path("/proj"))


class FakeConn(object):
    """pgwire.Connection の代役。script の要素を execute ごとに 1 つ使う。

    要素は (結果のリスト, 実行後のトランザクション状態) か、送出する例外。
    """

    def __init__(self, *script):
        self.script = list(script)
        self.executed = []
        self.transaction_status = "I"
        self.closed = False
        self.auth_method = "SCRAM-SHA-256"
        self.ssl_version = ""

    def execute(self, sql):
        self.executed.append(sql)
        step = self.script.pop(0) if self.script else ([], self.transaction_status)
        if isinstance(step, Exception):
            self.transaction_status = getattr(step, "status_after", "I")
            raise step
        results, self.transaction_status = step
        return results

    def close(self):
        self.closed = True


def result(columns, rows, command="SELECT"):
    r = pgwire.Result(columns, None)
    r.rows = rows
    r.command = command
    return r


class TestPostgresClient(unittest.TestCase):
    def client(self, conn):
        with mock.patch.object(db_mod, "connect_postgres", return_value=conn):
            return db_mod.PostgresClient(pg_settings())

    def test_query_returns_first_result_with_columns(self):
        conn = FakeConn(([result(None, [], "SET"), result(["a"], [["1"]]), result(["b"], [["2"]])], "I"))
        self.assertEqual(self.client(conn).query("SET …; SELECT 1; SELECT 2"), (["a"], [["1"]]))

    def test_query_with_params_is_refused(self):
        with self.assertRaises(db_mod.DbError):
            self.client(FakeConn()).query("SELECT ?", [1])

    def test_snapshot_keeps_server_text_untouched(self):
        conn = FakeConn(([result(["b", "ts", "n", "s", "x"], [["t", "2026-08-03 09:00:00.12", None, "ab  ", "\\x0a0b"]])], "I"))
        table = self.client(conn).snapshot({"name": "T"})
        self.assertEqual(table.rows, [["t", "2026-08-03 09:00:00.12", "(NULL)", "ab  ", "\\x0a0b"]])
        self.assertEqual(conn.executed, ["SELECT * FROM T"])

    def test_execute_script_commits_an_open_transaction(self):
        conn = FakeConn(([], "T"), ([], "I"))
        self.client(conn).execute_script("BEGIN; DELETE FROM t")
        self.assertEqual(conn.executed[-1], "COMMIT")

    def test_execute_script_rolls_back_and_reports_hint(self):
        error = pgwire.PgError("ERROR: duplicate key", {"C": "23505", "M": "duplicate key"})
        error.status_after = "E"
        conn = FakeConn(error, ([], "I"))
        with self.assertRaises(db_mod.DbError) as ctx:
            self.client(conn).execute_script("BEGIN; INSERT …")
        self.assertEqual(conn.executed[-1], "ROLLBACK")
        self.assertIn("一意制約", str(ctx.exception))

    def test_server_now_uses_clock_timestamp_and_parses_text(self):
        conn = FakeConn(([result(["now"], [["2026-10-22 10:11:12.5"]])], "I"))
        client = self.client(conn)
        self.assertEqual(db_mod.server_now(client), dt.datetime(2026, 10, 22, 10, 11, 12, 500000))
        self.assertIn("clock_timestamp()", conn.executed[0], "now() はトランザクション開始時刻で止まる")

    def test_batch_start_literal_keeps_microseconds(self):
        value = dt.datetime(2026, 8, 5, 0, 40, 47, 872123)
        self.assertEqual(db_mod.format_sql_datetime(value, "postgresql"), "2026-08-05T00:40:47.872123")
        self.assertEqual(db_mod.format_sql_datetime(value), "2026-08-05T00:40:47.872")

    def test_create_client_dispatches_by_type(self):
        with mock.patch.object(db_mod, "connect_postgres", return_value=FakeConn()):
            client = db_mod.create_client(pg_settings(), offline=False, case_id="TC001")
        self.assertIsInstance(client, db_mod.PostgresClient)

    def test_connect_passes_timeouts_and_session_settings(self):
        os.environ["UT_PG_PW"] = "pw"
        settings = pg_settings(server="pg01:6432", password_env="UT_PG_PW", query_timeout_sec=90,
                               timezone="Asia/Tokyo", search_path="app, public", sslmode="require",
                               sslrootcert="certs/ca.pem")
        with mock.patch.object(pgwire, "Connection") as factory:
            db_mod.connect_postgres(settings)
        args, kwargs = factory.call_args
        self.assertEqual(args, ("pg01", 6432, "app", "pw", "BATCH_DB"))
        self.assertEqual(kwargs["sslmode"], "require")
        self.assertEqual(Path(kwargs["sslrootcert"]), Path("/proj/certs/ca.pem"))
        self.assertEqual(kwargs["io_timeout"], 120.0)
        params = dict(kwargs["params"])
        self.assertEqual(params["statement_timeout"], "90000")
        self.assertEqual(params["TimeZone"], "Asia/Tokyo")
        self.assertEqual(params["search_path"], "app, public")

    def test_no_password_env_means_no_password(self):
        with mock.patch.object(pgwire, "Connection") as factory:
            db_mod.connect_postgres(pg_settings())
        self.assertIsNone(factory.call_args[0][3])


class TestPostgresLockHolder(unittest.TestCase):
    def test_lock_is_held_inside_a_transaction_and_rolled_back(self):
        conn = FakeConn(([], "T"), ([], "T"), ([], "I"))
        with mock.patch.object(db_mod, "connect_postgres", return_value=conn) as factory:
            with db_mod.LockHolder(pg_settings(), "LOCK TABLE t IN ACCESS EXCLUSIVE MODE"):
                self.assertEqual(conn.executed, ["BEGIN", "LOCK TABLE t IN ACCESS EXCLUSIVE MODE"])
        self.assertEqual(conn.executed[-1], "ROLLBACK")
        self.assertTrue(conn.closed)
        # batch の実行中にサーバ側のタイムアウトでロックが外れないように
        self.assertIn(("idle_in_transaction_session_timeout", "0"), factory.call_args[0][1])

    def test_lock_sql_that_commits_is_refused(self):
        """COMMIT でロックが外れたまま batch を動かすと、更新失敗の再現にならない。"""
        conn = FakeConn(([], "T"), ([], "I"))
        with mock.patch.object(db_mod, "connect_postgres", return_value=conn):
            with self.assertRaises(db_mod.DbError) as ctx:
                db_mod.LockHolder(pg_settings(), "LOCK TABLE t; COMMIT").__enter__()
        self.assertIn("トランザクションが残っていません", str(ctx.exception))
        self.assertTrue(conn.closed)

    def test_failing_lock_sql_is_reported(self):
        error = pgwire.PgError("ERROR: relation", {"C": "42P01"})
        conn = FakeConn(([], "T"), error)
        with mock.patch.object(db_mod, "connect_postgres", return_value=conn):
            with self.assertRaises(db_mod.DbError):
                db_mod.LockHolder(pg_settings(), "LOCK TABLE nope").__enter__()
        self.assertTrue(conn.closed)


class TestPgHints(unittest.TestCase):
    def hints(self, sqlstate, message="x"):
        return " ".join(db_mod.diagnose_pg_error(pgwire.PgError(message, {"C": sqlstate})))

    def test_sql_server_syntax_is_recognised(self):
        self.assertIn("SQL Server 用", self.hints("42601", 'ERROR: syntax error at or near "["'))
        self.assertIn("gen --overwrite", self.hints("42601", 'ERROR: syntax error at or near "GO"'))
        self.assertNotIn("SQL Server 用", self.hints("42601", 'ERROR: syntax error at or near "FORM"'))

    def test_common_states(self):
        self.assertIn("大文字小文字", self.hints("42703"))
        self.assertIn("identity=on", self.hints("428C9"))
        self.assertIn("password_env", self.hints("28P01"))
        self.assertIn("pg_hba.conf", self.hints("28000"))

    def test_connection_failures_without_sqlstate(self):
        self.assertIn("5432", " ".join(db_mod.diagnose_pg_error(pgwire.PgError("サーバに接続できません: pg01:5432"))))
        self.assertIn("sslrootcert", " ".join(db_mod.diagnose_pg_error(pgwire.PgError("SSL を確立できませんでした"))))


# =============================================================================
# 設定
# =============================================================================


class TestDatabaseSettings(unittest.TestCase):
    def load(self, database):
        root = Path(tempfile.mkdtemp(prefix="autotest_pgcfg_"))
        self.addCleanup(shutil.rmtree, str(root), True)
        path = root / "settings.yaml"
        path.write_text(yaml.safe_dump({"batch": {"exe_path": "x"}, "paths": {"log_dir": "./log"},
                                        "database": database}), encoding="utf-8")
        return load_settings(path, project_root=root, require_runtime=False)

    PG = {"type": "postgresql", "server": "pg01", "database": "BATCH_DB", "user": "app"}

    def pg(self, **extra):
        db = dict(self.PG)
        db.update(extra)
        return db

    def test_valid_postgresql_settings(self):
        s = self.load(self.pg(port=6432, sslmode="verify-full", sslrootcert="ca.pem", timezone="Asia/Tokyo",
                              search_path="app", password_env="PW", login_timeout_sec=5, query_timeout_sec=0))
        self.assertEqual(s.db_type, "postgresql")
        self.assertEqual(s.db_label(), "PostgreSQL pg01:6432")

    def test_sqlserver_is_the_default_and_label_is_unchanged(self):
        s = self.load({"server": "SQLSRV01,1433", "database": "D", "auth": "windows"})
        self.assertEqual(s.db_type, "sqlserver")
        self.assertEqual(s.db_label(), "SQLSRV01,1433")

    def assert_config_error(self, database, *parts):
        with self.assertRaises(ConfigError) as ctx:
            self.load(database)
        for part in parts:
            self.assertIn(part, str(ctx.exception))

    def test_unknown_type(self):
        self.assert_config_error(self.pg(type="postgres"), "sqlserver / postgresql")

    def test_sql_server_keys_under_postgresql_are_refused_with_reason(self):
        """encrypt: true を書いても PostgreSQL では効かない。黙って無視すると平文のまま繋がる。"""
        self.assert_config_error(self.pg(encrypt=True, driver="ODBC Driver 18 for SQL Server"),
                                 "encrypt", "sslmode", "driver")

    def test_postgresql_keys_under_sqlserver_are_refused(self):
        self.assert_config_error({"server": "s", "database": "d", "sslmode": "require"}, "sslmode")
        self.assert_config_error({"server": "s", "database": "d", "port": 1433}, "host,1433")

    def test_typo_in_sqlserver_key_is_refused(self):
        self.assert_config_error({"server": "s", "database": "d", "query_timeout": 30}, "query_timeout")

    def test_server_and_port_rules(self):
        self.assert_config_error(self.pg(server="pg01,5432"), "カンマ")
        self.assert_config_error(self.pg(server="pg01:5432", port=5432), "両方")
        self.assert_config_error(self.pg(port=70000), "ポート番号")
        self.assert_config_error(self.pg(server=":5432"), "ホスト名")

    def test_required_and_typed_values(self):
        db = self.pg()
        del db["user"]
        self.assert_config_error(db, "database.user")
        self.assert_config_error(self.pg(sslmode="on"), "sslmode")
        self.assert_config_error(self.pg(query_timeout_sec=-1), "query_timeout_sec")
        self.assert_config_error(self.pg(login_timeout_sec=True), "login_timeout_sec")
        self.assert_config_error({"server": "s", "database": "d", "auth": "kerberos"}, "sql / windows")

    def test_host_port_forms(self):
        self.assertEqual(pg_host_port({"server": "pg01"}), ("pg01", 5432))
        self.assertEqual(pg_host_port({"server": "pg01:6432"}), ("pg01", 6432))
        self.assertEqual(pg_host_port({"server": "pg01", "port": "6432"}), ("pg01", 6432))
        self.assertEqual(pg_host_port({"server": "[::1]:6432"}), ("::1", 6432))
        self.assertEqual(pg_host_port({"server": "::1"}), ("::1", 5432))


# =============================================================================
# dbcheck
# =============================================================================


class TestPostgresDbcheck(unittest.TestCase):
    def run_dbcheck(self, conn=None, reachable=True, error=None):
        settings = pg_settings(server="pg01:5432")
        identity = ["PostgreSQL 16.2 on x86_64-pc-linux-gnu, compiled by gcc", "BATCH_DB", "app",
                    "10.0.0.5", "5432", "Asia/Tokyo", '"$user", public']
        conn = conn or FakeConn(([result(["v"], [identity])], "I"))
        patches = [mock.patch.object(db_mod, "check_tcp_reachable",
                                     return_value=(True, "") if reachable else (False, "接続を拒否されました"))]
        if error is not None:
            patches.append(mock.patch.object(db_mod, "connect_postgres", side_effect=error))
        else:
            patches.append(mock.patch.object(db_mod, "connect_postgres", return_value=conn))
        buf = io.StringIO()
        with patches[0], patches[1], redirect_stdout(buf):
            code = cli._dbcheck(settings, timeout_sec=5)
        return code, buf.getvalue()

    @staticmethod
    def gui_scan(out):
        """gui._scan_line と同じ規則で、実接続先と相違の有無を拾う。"""
        found, mismatch = {}, False
        for line in out.splitlines():
            stripped = line.strip()
            for prefix, key in (("サーバ名   :", "server"), ("データベース:", "database")):
                if stripped.startswith(prefix):
                    found[key] = stripped[len(prefix):].strip()
            mismatch = mismatch or stripped.startswith("[警告] 接続先 DB が設定値")
        return found, mismatch

    def test_success_reports_actual_target_in_gui_format(self):
        code, out = self.run_dbcheck()
        self.assertEqual(code, 0, out)
        self.assertEqual(self.gui_scan(out), ({"server": "10.0.0.5:5432", "database": "BATCH_DB"}, False))
        self.assertIn("SCRAM-SHA-256", out)
        self.assertIn("Asia/Tokyo", out)

    def test_different_database_is_not_a_success(self):
        identity = ["PostgreSQL 16", "OTHER_DB", "app", "10.0.0.5", "5432", "UTC", "public"]
        code, out = self.run_dbcheck(conn=FakeConn(([result(["v"], [identity])], "I")))
        self.assertEqual(code, 1)
        self.assertEqual(self.gui_scan(out), ({"server": "10.0.0.5:5432", "database": "OTHER_DB"}, True))

    def test_unreachable_and_failed_connections(self):
        code, out = self.run_dbcheck(reachable=False)
        self.assertEqual(code, 1)
        self.assertIn("Test-NetConnection -ComputerName pg01 -Port 5432", out)
        code, out = self.run_dbcheck(error=db_mod.DbError("PostgreSQL への接続に失敗しました: 28P01"))
        self.assertEqual(code, 1)
        self.assertIn("[NG] 接続に失敗しました", out)

    def test_dbcheck_runs_even_when_every_case_is_disabled(self):
        """gen 直後は全ケースが enabled: false。それでも接続確認はできること。"""
        root = Path(tempfile.mkdtemp(prefix="autotest_pgcheck_"))
        self.addCleanup(shutil.rmtree, str(root), True)
        (root / "config").mkdir()
        (root / "cases").mkdir()
        (root / "config" / "settings.yaml").write_text(yaml.safe_dump({
            "batch": {"exe_path": "x"}, "paths": {"log_dir": "./log"},
            "database": {"type": "postgresql", "server": "pg01", "database": "D", "user": "u"}}),
            encoding="utf-8")
        (root / "cases" / "TC001.yaml").write_text("id: TC001\nname: n\nenabled: false\n", encoding="utf-8")
        out, err = io.StringIO(), io.StringIO()
        with mock.patch.object(db_mod, "check_tcp_reachable", return_value=(False, "拒否")), \
                redirect_stdout(out), redirect_stderr(err):
            code = cli.main(["dbcheck", "--config", str(root / "config" / "settings.yaml")])
        self.assertEqual(code, 1, err.getvalue())
        self.assertIn("PostgreSQL 接続確認", out.getvalue())


# =============================================================================
# gen（テストケース md → seed.sql / スナップショット SQL）
# =============================================================================


FRONT = "---\n出力先: grp\n---\n# タイトル\n\n"
CASE = ("## TC201 正常系：PostgreSQL\n"
        "- 種別: 正常系\n"
        "- 目的: 投入できること\n"
        "- 終了コード: 0\n")


class PgGenCase(unittest.TestCase):
    DATABASE = {"type": "postgresql", "server": "pg01", "database": "D", "user": "u"}

    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="autotest_pggen_")).resolve()
        self.addCleanup(shutil.rmtree, str(self.root), True)
        (self.root / "config").mkdir()
        (self.root / "config" / "settings.yaml").write_text(yaml.safe_dump({
            "env": {"sandbox": True}, "database": self.DATABASE,
            "batch": {"exe_path": "{python}", "working_dir": "."},
            "paths": {n: "sandbox/" + n for n in ("input_dir", "processed_dir", "error_dir", "output_dir", "log_dir")},
        }), encoding="utf-8")
        self.md = self.root / "docs" / "cases.md"
        self.md.parent.mkdir(parents=True)

    def gen(self, body):
        self.md.write_text(FRONT + body, encoding="utf-8")
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = cli.main(["gen", "--md", str(self.md), "--config", str(self.root / "config" / "settings.yaml")])
        return code, out.getvalue(), err.getvalue()

    def seed_sql(self):
        return (self.root / "cases" / "grp" / "TC201" / "seed.sql").read_text(encoding="utf-8")


class TestPostgresGen(PgGenCase):
    def test_seed_uses_postgresql_syntax(self):
        code, out, err = self.gen(CASE + "\n```csv seed=public.t_order key=order_id\n"
                                         "order_id,memo,path,note\n1,ab  ,C:\\in,(NULL)\n2,O'Brien,x,\n```\n")
        self.assertEqual(code, 0, err)
        sql = self.seed_sql()
        self.assertIn("PostgreSQL 用", sql.splitlines()[0])
        self.assertIn("""DELETE FROM public.t_order WHERE "order_id" IN ('1', '2');""", sql)
        self.assertIn("""INSERT INTO public.t_order ("order_id", "memo", "path", "note") VALUES""", sql)
        # N'...' は character 型になり、varchar / text へ入れると末尾の空白が落ちる
        self.assertIn("('1', 'ab  ', E'C:\\\\in', NULL)", sql)
        self.assertIn("('2', 'O''Brien', 'x', '')", sql)
        self.assertNotIn("N'", sql)
        self.assertNotIn("[order_id]", sql)
        self.assertNotIn("\nGO", sql)
        self.assertNotIn("OVERRIDING SYSTEM VALUE", sql, "identity=on のときだけ")
        self.assertIn("pg_get_serial_sequence('public.t_order', col)", sql, "採番列のシーケンスを進める")
        self.assertIn("数値列・日付列ではエラーになります", out)
        self.assertIn("DB の種類       : PostgreSQL", out)

    def test_identity_on_overrides_generated_always(self):
        code, _out, err = self.gen(CASE + "\n```csv seed=t_audit delete=all identity=on\naudit_id,msg\n1,a\n```\n")
        self.assertEqual(code, 0, err)
        sql = self.seed_sql()
        self.assertIn("DELETE FROM t_audit;", sql)
        self.assertIn('INSERT INTO t_audit ("audit_id", "msg") OVERRIDING SYSTEM VALUE VALUES', sql)
        self.assertNotIn("IDENTITY_INSERT", sql)

    def test_expect_db_select_quotes_columns(self):
        code, _out, err = self.gen(CASE + "\n```csv expect_db=t_order key=order_id ignore=updated_at\n"
                                          "order_id,status,updated_at\n1,完了,x\n```\n")
        self.assertEqual(code, 0, err)
        data = yaml.safe_load((self.root / "cases" / "grp" / "TC201.yaml").read_text(encoding="utf-8"))
        self.assertEqual(data["snapshot"]["tables"], [{
            "name": "t_order",
            "sql": 'SELECT "order_id", "status", "updated_at" FROM t_order ORDER BY "order_id"'}])

    def test_sql_server_table_syntax_is_rejected(self):
        code, _out, err = self.gen(CASE + "\n```csv seed=[dbo].[T] key=ID\nID\n1\n```\n")
        self.assertEqual(code, 2)
        self.assertIn("[ ]", err)


class TestGenWithoutReadableSettings(PgGenCase):
    DATABASE = {"type": "oracle", "server": "s", "database": "d"}

    def test_only_sql_blocks_need_the_database_type(self):
        code, _out, err = self.gen(CASE + "- ファイル: `output_dir/*.csv` = 1\n")
        self.assertEqual(code, 0, err)
        code, _out, err = self.gen(CASE + "\n```csv seed=T key=ID\nID\n1\n```\n")
        self.assertEqual(code, 2)
        self.assertIn("database.type", err)


if __name__ == "__main__":
    unittest.main()
