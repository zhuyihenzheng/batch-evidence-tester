# -*- coding: utf-8 -*-
"""接続先 DB の確認（autotest dbcheck）の回帰テスト。

「繋がったか」だけを見て合格にすると、設定ミスで別の DB に繋がったまま
試験が通ってしまう。結果は正常に見えるのに見ていた DB が違う——
偽 OK の中でも最も気づきにくい形なので、ここは終了コードで区別する。

  python -m unittest discover -s tests -v
"""

import io
import sys
import types
import unittest
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from autotest import cli, db as db_mod  # noqa: E402
from autotest.config import Settings  # noqa: E402


class FakeCursor(object):
    def __init__(self, row):
        self._row = row

    def execute(self, sql):
        return None

    def fetchone(self):
        return self._row

    def close(self):
        return None


class FakeConnection(object):
    def __init__(self, row):
        self._row = row
        self.closed = False

    def cursor(self):
        return FakeCursor(self._row)

    def close(self):
        self.closed = True


class DbCheckCase(unittest.TestCase):
    """pyodbc と各種診断を差し替えて _dbcheck だけを動かす。"""

    CONFIGURED_DB = "BATCH_DB"

    def setUp(self):
        self.settings = Settings(
            raw={
                "batch": {"exe_path": "dummy.exe"},
                "paths": {"log_dir": "./log"},
                "database": {
                    "server": "SQLSRV01,1433",
                    "database": self.CONFIGURED_DB,
                    "driver": "ODBC Driver 18 for SQL Server",
                    "auth": "sql",
                    "user": "sa",
                    "password_env": "AUTOTEST_TEST_PW",
                },
            },
            source=Path("settings.yaml"),
            project_root=Path("."),
        )
        self._saved = {}
        for name in ("list_installed_drivers", "diagnose_password_env",
                     "parse_server", "check_tcp_reachable", "build_connection_string"):
            self._saved[name] = getattr(db_mod, name)
        db_mod.list_installed_drivers = lambda: ["ODBC Driver 18 for SQL Server"]
        db_mod.diagnose_password_env = lambda name, value: ["[OK] 設定されています"]
        db_mod.parse_server = lambda s: ("SQLSRV01", 1433, None)
        db_mod.check_tcp_reachable = lambda host, port, timeout_sec=5: (True, "")
        db_mod.build_connection_string = lambda s, timeout_sec=None: ("CONN", "CONN(表示用)")
        self._saved_pyodbc = sys.modules.get("pyodbc")

    def tearDown(self):
        for name, func in self._saved.items():
            setattr(db_mod, name, func)
        if self._saved_pyodbc is None:
            sys.modules.pop("pyodbc", None)
        else:
            sys.modules["pyodbc"] = self._saved_pyodbc

    def run_dbcheck(self, actual_db):
        """実際に繋がった DB 名を指定して dbcheck を動かす。(終了コード, 出力)"""
        fake = types.ModuleType("pyodbc")
        fake.connect = lambda conn_str, timeout=None: FakeConnection(
            ("Microsoft SQL Server 2019", actual_db, "sa", "SQLSRV01"))
        sys.modules["pyodbc"] = fake

        buf = io.StringIO()
        with redirect_stdout(buf):
            code = cli._dbcheck(self.settings, timeout_sec=5)
        return code, buf.getvalue()


class TestDbCheckVerdict(DbCheckCase):
    def test_matching_database_succeeds(self):
        code, out = self.run_dbcheck(self.CONFIGURED_DB)
        self.assertEqual(code, 0)
        self.assertIn("接続確認 OK", out)

    def test_connecting_to_a_different_database_is_not_a_success(self):
        """繋がっても接続先が設定と違えば合格にしないこと。

        ここが 0 を返すと、別の DB を見た試験結果が「事前確認は通った」
        という記録つきで残ってしまう。
        """
        code, out = self.run_dbcheck("BATCH_DB_OLD")
        self.assertEqual(code, 1, "接続先が違うのに成功扱いになっている")
        self.assertIn("接続先 DB が設定値", out)
        self.assertNotIn("接続確認 OK", out)

    def test_actual_target_is_reported_not_just_the_setting(self):
        """設定値の読み上げではなく、DB に問い合わせた実際の接続先を出すこと。"""
        _, out = self.run_dbcheck("BATCH_DB_OLD")
        self.assertIn("実際の接続先:", out)
        self.assertIn("BATCH_DB_OLD", out)
        self.assertIn("SQLSRV01", out)

    def test_mismatch_message_explains_the_consequence(self):
        """非開発者が読んで何が起きるか分かること。"""
        _, out = self.run_dbcheck("BATCH_DB_OLD")
        self.assertIn("別の DB を見た結果が証跡として残ります", out)


class TestDbCheckOutputIsParseableByGui(DbCheckCase):
    """操作画面は dbcheck の出力から実接続先を拾う。書式を変えたら画面が壊れる。"""

    def test_gui_can_extract_actual_server_and_database(self):
        _, out = self.run_dbcheck("BATCH_DB_OLD")

        # gui._scan_line と同じ規則で拾えること（tkinter が無い環境でも
        # 検査できるよう、抽出規則そのものをここで再現している）
        found = {}
        mismatch = False
        for line in out.splitlines():
            stripped = line.strip()
            for prefix, key in (("サーバ名   :", "server"), ("データベース:", "database")):
                if stripped.startswith(prefix):
                    found[key] = stripped[len(prefix):].strip()
            if stripped.startswith("[警告] 接続先 DB が設定値"):
                mismatch = True

        self.assertEqual(found.get("server"), "SQLSRV01")
        self.assertEqual(found.get("database"), "BATCH_DB_OLD")
        self.assertTrue(mismatch, "画面が相違を検出できない")


if __name__ == "__main__":
    unittest.main()
