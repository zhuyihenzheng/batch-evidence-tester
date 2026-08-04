"""コマンドラインエントリポイント。

  python -m autotest run                    全ケース実行
  python -m autotest run --case TC001       ケース指定
  python -m autotest run --tag 正常系        タグ指定
  python -m autotest run --offline          DB/exe 無しでパイプライン検証
  python -m autotest validate               設定とケース定義の妥当性チェックのみ
"""

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from .config import ConfigError, load_cases, load_settings
from .excel import build_workbook
from .models import OK, RunResult
from .orchestrator import CaseRunner
from .runner import _resolve_exe

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


def _prepare_console() -> None:
    """コンソール出力で UnicodeEncodeError を出さないようにする。

    日本語 Windows の既定コンソールは cp932 で、日本語は出せるが一部記号が出せない。
    落とさず「?」に置換して続行させる（`chcp 65001` すればそのまま出る）。
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure:
            try:
                reconfigure(errors="replace")
            except (ValueError, OSError):
                pass


def _project_root_for(config_path: Path) -> Path:
    """プロジェクトルートを決める。

    設定ファイルは <root>/config/settings.yaml に置く規約なので、その 2 階層上を
    ルートとみなす。規約から外れた場所を指された場合はパッケージ位置から推測する。
    """
    candidate = config_path.resolve().parent.parent
    return candidate if (candidate / "cases").is_dir() else PROJECT_ROOT


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="autotest", description="batch(.exe) 自動テスト実行ツール")
    parser.add_argument("command", choices=["run", "validate", "list", "dbcheck"],
                        help="実行するコマンド（dbcheck: SQL Server へ実際に接続できるか確認）")
    parser.add_argument("--config", default=str(PROJECT_ROOT / "config" / "settings.yaml"), help="設定ファイル")
    parser.add_argument("--cases-dir", default=None, help="ケース定義フォルダ（既定: <プロジェクトルート>/cases）")
    parser.add_argument("--case", action="append", dest="cases", help="実行するケースID（複数指定可）")
    parser.add_argument("--tag", action="append", dest="tags", help="実行するタグ（複数指定可）")
    parser.add_argument("--out", default=None, help="Excel 出力先（既定: settings.excel.output_dir）")
    parser.add_argument("--offline", action="store_true", help="SQL Server の代わりに fixtures/ の CSV を使う")
    parser.add_argument("--dry-run", action="store_true",
                        help=".exe も DB 接続も行わず、流れと設定だけを確認する")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="各工程の開始を逐次表示する（処理が止まる箇所の切り分け用）")
    # 注: NG が出ても後続ケースは常に実行する（全 NG を一度に把握するため）。
    #     以前あった --keep-going は「定義だけあって未使用」だったため削除した。
    return parser


def _dbcheck(settings, timeout_sec: int = 10) -> int:
    """SQL Server へ実際に接続できるかを確認する。

    validate は設定の書式しか見ないため、「本当に繋がるか」は別途確認が要る。
    失敗時は原因の切り分けに必要な情報（導入済みドライバ一覧など）まで出す。
    """
    from . import db as db_mod

    print("=" * 66)
    print(" SQL Server 接続確認")
    print("=" * 66)

    db = settings.database
    auth = str(db.get("auth", "sql")).lower()
    print("  接続先     : %s / %s" % (db.get("server"), db.get("database")))
    print("  認証方式   : %s" % ("Windows 認証" if auth == "windows" else "SQL Server 認証（ユーザー: %s）" % db.get("user")))
    print("  ドライバ   : %s" % db.get("driver"))

    # --- pyodbc の有無 ------------------------------------------------------
    try:
        import pyodbc  # noqa: F401,PLC0415
    except ImportError:
        print("\n[NG] pyodbc が入っていません。")
        print("     Anaconda      : conda install pyodbc")
        print("     通常の Python : pip install pyodbc")
        return 1

    # --- 導入済みドライバ ----------------------------------------------------
    drivers = db_mod.list_installed_drivers()
    print("\n  この端末に入っている ODBC ドライバ:")
    if drivers:
        for d in drivers:
            mark = "  <- 設定値と一致" if d == db.get("driver") else ""
            print("    - %s%s" % (d, mark))
        if db.get("driver") not in drivers:
            print("\n[警告] 設定値 '%s' が一覧にありません。上の中から選んでください。" % db.get("driver"))
    else:
        print("    （1 つも見つかりません）")
        print("\n[NG] ODBC ドライバが未導入です。")
        print("     Microsoft ODBC Driver for SQL Server を入れてください。")
        return 1

    # --- パスワード環境変数 --------------------------------------------------
    if auth != "windows":
        env_name = str(db.get("password_env", ""))
        print("\n  パスワード環境変数 %s:" % env_name)
        for line in db_mod.diagnose_password_env(env_name, os.environ.get(env_name)):
            print("    %s" % line)

    # --- TCP 到達性 ---------------------------------------------------------
    # ODBC のエラーは原因が読み取りにくいので、ネットワークの問題かどうかを
    # ここで切り分ける。ローカルで完結する確認（ドライバ・環境変数）を先に
    # 済ませてから行う。ネットワークで止まっても、手前の情報は出し切るため。
    host, port, instance = db_mod.parse_server(str(db.get("server", "")))
    if instance and port is None:
        print("\n  名前付きインスタンス '%s' はポートが動的です。" % instance)
        print("  SQL Server Browser (UDP 1434) 経由で解決されるため、TCP 事前確認は省略します。")
        print("  接続が不安定な場合は固定ポートを割り当て、server: \"%s,<ポート>\" と書く方が確実です。" % host)
    elif host and port:
        print("\n  TCP 到達確認 : %s:%d" % (host, port))
        ok, reason = db_mod.check_tcp_reachable(host, port, timeout_sec=min(5, timeout_sec))
        if ok:
            print("    [OK] ポートまで到達できています")
        else:
            print("    [NG] %s" % reason)
            print("\n  ネットワーク層で止まっています。次を確認してください:")
            print("    - サーバ名 / ポートの誤り（SQL Server はコロンではなくカンマ区切り: host,1433）")
            print("    - セキュリティグループ / ファイアウォールで %d 番が開いているか" % port)
            print("    - SQL Server 側で TCP/IP プロトコルが有効か")
            print("    - Windows 側からは次で追加確認できます:")
            print("        Test-NetConnection -ComputerName %s -Port %d" % (host, port))
            return 1

    # --- 接続文字列 ----------------------------------------------------------
    try:
        conn_str, shown = db_mod.build_connection_string(settings)
    except ConfigError as exc:
        print("\n[NG] %s" % exc)
        return 1
    print("\n  接続文字列 : %s" % shown)

    # --- 実接続 --------------------------------------------------------------
    print("\n  接続中（タイムアウト %d 秒）..." % timeout_sec)
    started = datetime.now()
    try:
        conn = pyodbc.connect(conn_str, timeout=timeout_sec)
    except Exception as exc:
        elapsed = (datetime.now() - started).total_seconds()
        print("\n[NG] 接続に失敗しました（%.1f 秒）" % elapsed)
        print("     %s" % str(exc).replace("\n", "\n     "))
        hints = db_mod.diagnose_connection_error(exc)
        if hints:
            print("\n  確認してください:")
            for h in hints:
                print("    %s" % h)
        return 1

    elapsed = (datetime.now() - started).total_seconds()
    print("[OK] 接続成功（%.2f 秒）" % elapsed)

    # --- 接続先の実体を確認（設定ミスで別 DB に繋がっていないか）--------------
    try:
        cur = conn.cursor()
        cur.execute("SELECT @@VERSION, DB_NAME(), SUSER_NAME(), @@SERVERNAME")
        version, dbname, login, servername = cur.fetchone()
        print("\n  実際の接続先:")
        print("    サーバ名   : %s" % servername)
        print("    データベース: %s" % dbname)
        print("    ログイン   : %s" % login)
        print("    バージョン : %s" % str(version).splitlines()[0].strip())
        if dbname != db.get("database"):
            print("\n[警告] 接続先 DB が設定値 '%s' と異なります。" % db.get("database"))
        cur.close()
    except Exception as exc:
        print("\n[警告] 接続はできましたが情報取得に失敗しました: %s" % exc)
    finally:
        conn.close()

    print("\n" + "=" * 66)
    print(" 接続確認 OK。次は python -m autotest validate で設定全体を確認してください。")
    print("=" * 66)
    return 0


def _is_under(path: Path, parent: Path) -> bool:
    """path が parent の配下か。Path.is_relative_to は 3.9+ のため自前で判定する。"""
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def _validate(settings, cases, args, project_root: Path) -> int:
    """設定とケース定義の妥当性チェック。

    問題を発見したら終了コード 1 を返す。以前は問題を表示しつつ 0 を
    返していたため、CI や setup スクリプトが「設定は有効」と誤解していた。
    """
    from .orchestrator import preflight_case  # 遅延 import（循環回避）

    errors: List[str] = []
    warnings: List[str] = []

    print(f"設定ファイル : {args.config}")
    print(f"ケース定義   : {len(cases)} 件")

    # --- paths（run 時に自動作成されるため「未作成」は警告扱い）--------------
    raw_paths = settings.raw.get("paths") or {}
    # サンドボックス設定はプロジェクト配下の相対パスを意図的に使うため、
    # 「相対パス警告」の対象から外す（env.sandbox: true で明示する）
    is_sandbox = bool(settings.env.get("sandbox", False))
    for alias, path in settings.path_aliases.items():
        mark = "存在" if path.is_dir() else "未作成（run 時に自動作成）"
        if not path.is_dir():
            warnings.append(f"paths.{alias} は未作成です: {path}")
        print(f"  paths.{alias:<16} {path}  ({mark})")

        # 相対パスはツール自身のフォルダ配下に解決される。テスト対象 batch の
        # フォルダを指すつもりで相対で書くと、まったく別の場所を見てしまうため警告する
        raw = str(raw_paths.get(alias, ""))
        if not is_sandbox and raw and not Path(raw).is_absolute() and _is_under(path, project_root):
            warnings.append(
                f"paths.{alias} が相対パスのため、ツール自身のフォルダ配下に解決されています。\n"
                f"          設定値: {raw}\n"
                f"          解決先: {path}\n"
                f"          テスト対象 batch のフォルダを指すなら、ドライブ文字付きの"
                f"絶対パス（例 C:/app/batch/in）で書いてください。"
            )

    # --- 既定 batch ---------------------------------------------------------
    exe = _resolve_exe(str(settings.batch["exe_path"]), project_root)
    exe_ok = exe.exists()
    if not exe_ok:
        errors.append(f"batch.exe_path が見つかりません: {exe}")
    print(f"  batch.exe_path   {exe}  ({'存在' if exe_ok else '★見つかりません'})")

    # --- 名前付き batch -------------------------------------------------------
    for name in settings.batches:
        profile = settings.batch_profile(name)
        bexe = _resolve_exe(str(profile.get("exe_path", "")), project_root)
        bexe_ok = bexe.exists()
        if not bexe_ok:
            errors.append(f"batches.{name}.exe_path が見つかりません: {bexe}")
        print(f"  batches.{name:<12} {bexe}  ({'存在' if bexe_ok else '★見つかりません'})")

    print(f"  database         {settings.database.get('server')} / {settings.database.get('database')}")

    # --- DB パスワード環境変数（offline 運用もあるため警告扱い）--------------
    env_name = settings.database.get("password_env")
    if env_name and str(settings.database.get("auth", "sql")).lower() != "windows":
        if os.environ.get(env_name) is None:
            warnings.append(f"DB パスワード環境変数 {env_name} が未設定です（--offline のみなら不要）")

    # --- ケースごとの preflight（実行時と同じ検査）---------------------------
    print()
    for case in cases:
        problems = preflight_case(settings, case)
        status = "OK" if not problems else "NG"
        print(f"  [{status}] {case.case_id}  {case.name}")
        for p in problems:
            print(f"        - {p}")
            errors.append(f"{case.case_id}: {p}")

    print()
    for w in warnings:
        print(f"[警告] {w}")
    if errors:
        print(f"[NG] 問題 {len(errors)} 件。上記を解消してから実行してください。")
        return 1
    print("[OK] 検証を通過しました。")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    _prepare_console()
    args = build_parser().parse_args(argv)

    config_path = Path(args.config)
    project_root = _project_root_for(config_path)
    cases_dir = Path(args.cases_dir) if args.cases_dir else project_root / "cases"

    try:
        settings = load_settings(config_path, project_root=project_root)
        cases = load_cases(cases_dir, only=args.cases, tags=args.tags)
    except ConfigError as exc:
        print(f"[設定エラー] {exc}", file=sys.stderr)
        return 2

    if args.command == "list":
        print(f"{'ケースID':<16}{'区分':<16}ケース名")
        print("-" * 72)
        for c in cases:
            print(f"{c.case_id:<16}{'/'.join(c.tags):<16}{c.name}")
        return 0

    if args.command == "dbcheck":
        return _dbcheck(settings, timeout_sec=int(settings.database.get("login_timeout_sec", 15)))

    if args.command == "validate":
        return _validate(settings, cases, args, project_root)

    # --- run --------------------------------------------------------------
    started = datetime.now()
    # ミリ秒まで含める。秒精度だと同一秒に 2 回起動した場合に
    # 証跡フォルダと Excel が同名になり、互いに上書きされる
    run_id = started.strftime("%Y%m%d_%H%M%S") + "_%03d" % (started.microsecond // 1000)
    out_dir = Path(args.out).resolve() if args.out else settings.resolve_dir(str(settings.excel.get("output_dir", "./output")))
    run_dir = out_dir / run_id

    run = RunResult(
        run_id=run_id,
        started_at=started,
        env_name=str(settings.env.get("name", "")),
        tester=settings.tester,
        exe_path=str(settings.batch.get("exe_path", "")),
        db_server=str(settings.database.get("server", "")),
        db_name=str(settings.database.get("database", "")),
    )

    mode_note = " [offline]" if args.offline else ""
    mode_note += " [dry-run]" if args.dry_run else ""
    print(f"=== 自動テスト開始{mode_note}  run_id={run_id}  ケース数={len(cases)} ===")

    # 工程ログ。処理が止まったとき「どこで止まっているか」を後から追えるようにする。
    # --verbose なら画面にも出す。ファイルには常に記録する。
    run_dir.mkdir(parents=True, exist_ok=True)
    log_path = run_dir / "run.log"
    log_file = log_path.open("w", encoding="utf-8")

    def write_log(message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        log_file.write(f"{stamp}  {message}\n")
        log_file.flush()  # ハング時に読めないと意味がないので都度フラッシュ

    write_log(f"開始 run_id={run_id} config={config_path} offline={args.offline} dry_run={args.dry_run}")
    print(f"実行ログ : {log_path}")

    try:
        runner = CaseRunner(settings, run_dir, offline=args.offline, dry_run=args.dry_run)
        for i, case in enumerate(cases, start=1):
            head = f"[{i}/{len(cases)}] {case.case_id} {case.name}"
            print(head + " ... ", end="" if not args.verbose else "\n", flush=True)
            write_log(f"--- {case.case_id} 開始")

            def on_step(step: str, _case_id=case.case_id) -> None:
                write_log(f"    {_case_id}: {step}")
                if args.verbose:
                    print(f"      > {step}", flush=True)

            runner._progress = on_step
            result = runner.run(case)
            run.cases.append(result)

            detail = f" ({result.ng_count} 件 NG)" if result.ng_count else ""
            print(f"{'      => ' if args.verbose else ''}{result.verdict}{detail}")
            write_log(f"--- {case.case_id} 終了: {result.verdict}{detail}")
            if result.fatal_error:
                print(f"    実行時エラー: {result.fatal_error.splitlines()[0]}")
                write_log(f"    実行時エラー: {result.fatal_error}")
    finally:
        log_file.close()

    run.finished_at = datetime.now()

    file_name = str(settings.excel.get("file_name_format", "TestEvidence_{run_id}.xlsx")).format(run_id=run_id)
    excel_path = build_workbook(run, out_dir / file_name, settings.excel)

    print("=" * 60)
    if args.dry_run:
        # dry-run は判定していないので、成否ではなく「流れが通ったか」を返す
        failed = [c for c in run.cases if c.fatal_error]
        print(f"dry-run 完了 : {len(run.cases) - len(failed)} / {len(run.cases)} ケースが前処理まで通過")
        for c in failed:
            print(f"  [失敗] {c.case_id}: {c.fatal_error.splitlines()[0]}")
        print("※ batch も DB も実行していないため判定は行っていません。")
        print(f"証跡Excel: {excel_path}")
        print(f"実行ログ : {run_dir / 'run.log'}")
        return 1 if failed else 0

    print(f"総合判定 : {run.verdict}   OK {run.ok_count} / NG {run.ng_count}")
    if run.verdict == "SKIP":
        print("※ 実際に判定されたケースがありません（全件 SKIP）。終了コードは 1 になります。")
    print(f"証跡Excel: {excel_path}")
    print(f"証跡画像 : {run_dir / 'evidence'}")
    print(f"回収物   : {run_dir / 'artifacts'}")
    return 0 if run.verdict == OK else 1


if __name__ == "__main__":
    sys.exit(main())
