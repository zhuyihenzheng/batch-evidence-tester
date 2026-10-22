# -*- coding: utf-8 -*-
"""実際の PostgreSQL サーバに対する確認。環境変数 AUTOTEST_PG_TEST_URL があるときだけ動く。

  AUTOTEST_PG_TEST_URL=postgresql://user:password@127.0.0.1:5432/dbname \\
      python -m unittest tests.test_postgresql_live -v

表は autotest_live_<pid>_* の名前で作り、終わったら消す。CI や Windows の
通常のテスト実行ではスキップされる。
"""

import datetime as dt
import os
import sys
import unittest
from pathlib import Path

from urllib.parse import unquote, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import casegen, pgwire  # noqa: E402
from autotest import db as db_mod  # noqa: E402
from autotest.config import Settings  # noqa: E402

URL = os.environ.get("AUTOTEST_PG_TEST_URL")


def live_settings():
    url = urlparse(URL)
    os.environ["AUTOTEST_PG_LIVE_PW"] = unquote(url.password or "")
    database = {"type": "postgresql", "server": "%s:%d" % (url.hostname, url.port or 5432),
                "database": url.path.lstrip("/"), "user": unquote(url.username or ""),
                "password_env": "AUTOTEST_PG_LIVE_PW", "timezone": "Asia/Tokyo"}
    return Settings(raw={"batch": {"exe_path": "x"}, "paths": {"log_dir": "./log"}, "database": database},
                    source=Path("settings.yaml"), project_root=Path("."))


@unittest.skipUnless(URL, "AUTOTEST_PG_TEST_URL が未設定（実サーバでの確認はスキップ）")
class TestLivePostgres(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.settings = live_settings()
        cls.prefix = "autotest_live_%d" % os.getpid()
        cls.client = db_mod.PostgresClient(cls.settings)

    @classmethod
    def tearDownClass(cls):
        cls.client.conn.execute("DROP TABLE IF EXISTS %s_values, %s_serial, %s_identity"
                                % (cls.prefix, cls.prefix, cls.prefix))
        cls.client.close()

    def test_values_are_returned_as_the_server_prints_them(self):
        _cols, rows = self.client.query(
            "SELECT NULL::text, ''::text, 'ab  '::varchar, true, 15000::numeric(12,2),"
            " '2026-08-03 09:00:00.12'::timestamp, '\\x0a0b'::bytea, '東京'::text")
        self.assertEqual(rows, [[None, "", "ab  ", "t", "15000.00", "2026-08-03 09:00:00.12", "\\x0a0b", "東京"]])

    def test_generated_seed_round_trips_exactly(self):
        table = self.prefix + "_values"
        self.client.execute_script("DROP TABLE IF EXISTS %s; CREATE TABLE %s (id integer PRIMARY KEY,"
                                   " memo varchar(20), path text, note text, qty integer)" % (table, table))
        header = ["id", "memo", "path", "note", "qty"]
        rows = [["1", "ab  ", "C:\\in\\x", "(NULL)", "10"], ["2", "O'Brien", "", "東京", "(NULL)"]]
        sql = casegen._seed_sql(table, header, rows, ["id"], "key", False, casegen.PostgresDialect)
        self.client.execute_script("\n".join(sql))
        _cols, got = self.client.query('SELECT "id", "memo", "path", "note", "qty" FROM %s ORDER BY "id"' % table)
        self.assertEqual([[db_mod.to_text(v) for v in row] for row in got], rows)

    def test_seeded_serial_and_identity_do_not_collide_with_the_next_insert(self):
        for suffix, ddl, identity in (("serial", "id serial PRIMARY KEY", False),
                                      ("identity", "id integer GENERATED ALWAYS AS IDENTITY PRIMARY KEY", True)):
            table = "%s_%s" % (self.prefix, suffix)
            self.client.execute_script("DROP TABLE IF EXISTS %s; CREATE TABLE %s (%s, v text)" % (table, table, ddl))
            sql = casegen._seed_sql(table, ["id", "v"], [["1", "a"], ["2", "b"]], [], "all", identity,
                                    casegen.PostgresDialect)
            self.client.execute_script("\n".join(sql))
            # batch が採番に任せて追加する場面。シーケンスが進んでいないと id=1 で重複する
            self.client.execute_script("INSERT INTO %s (v) VALUES ('batch')" % table)
            _cols, got = self.client.query("SELECT max(id) FROM %s" % table)
            self.assertGreater(int(got[0][0]), 2, suffix)

    def test_lock_holder_blocks_other_sessions_until_released(self):
        table = self.prefix + "_values"
        self.client.execute_script("CREATE TABLE IF NOT EXISTS %s (id integer PRIMARY KEY, memo varchar(20),"
                                   " path text, note text, qty integer)" % table)

        def update():
            other = db_mod.connect_postgres(self.settings)
            try:
                other.execute("SET lock_timeout = '300ms'; UPDATE %s SET qty = qty" % table)
                return "ok"
            except pgwire.PgError as exc:
                return exc.sqlstate
            finally:
                other.close()

        with db_mod.LockHolder(self.settings, "LOCK TABLE %s IN ACCESS EXCLUSIVE MODE" % table):
            self.assertEqual(update(), "55P03")
        self.assertEqual(update(), "ok")

    def test_server_now_is_the_current_time(self):
        now = db_mod.server_now(self.client)
        self.assertIsInstance(now, dt.datetime)
        self.assertLess(abs((now - dt.datetime.now()).total_seconds()), 24 * 3600)


if __name__ == "__main__":
    unittest.main()
