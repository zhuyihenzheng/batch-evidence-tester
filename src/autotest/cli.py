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

from . import result_store
from .config import ConfigError, load_cases, load_settings
from .excel import build_workbook
from .models import OK, REVIEW, RunResult
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


def _filter_description(args) -> str:
    """証跡に残す絞り込み条件の説明。

    一部だけ実行した結果を「全体が合格した」と誤読させないため、
    どう絞り込んだかを証跡側に必ず記録する。
    """
    parts = []
    if args.cases:
        parts.append("ケース指定: " + ", ".join(args.cases))
    if args.tags:
        parts.append("タグ指定: " + ", ".join(args.tags))
    if not parts:
        return ""
    return " / ".join(parts)


def _filter_slug(args) -> str:
    """ファイル名に使える形の絞り込み識別子。絞り込みが無ければ空文字。"""
    parts = []
    if args.tags:
        parts.append("tag-" + "-".join(args.tags))
    if args.cases:
        parts.append("case-" + "-".join(args.cases))
    slug = "_".join(parts)
    return "".join(ch for ch in slug if ch.isalnum() or ch in "-_ぁ-んァ-ヶ一-龠") or "all"


def _default_config() -> Path:
    """既定の設定ファイルを決める。

    config/settings.local.yaml があればそれを優先する。こちらは .gitignore
    済みなので、各自の環境値（実パス・DB 接続先）を書いても git pull で
    衝突しない。settings.yaml はリポジトリ側が更新し続けるテンプレート。
    """
    local = PROJECT_ROOT / "config" / "settings.local.yaml"
    if local.exists():
        return local
    return PROJECT_ROOT / "config" / "settings.yaml"


def _project_root_for(config_path: Path) -> Path:
    """プロジェクトルートを決める。

    設定ファイルは <root>/config/settings.yaml に置く規約なので、その 2 階層上を
    ルートとみなす。規約から外れた場所を指された場合はパッケージ位置から推測する。
    """
    candidate = config_path.resolve().parent.parent
    return candidate if (candidate / "cases").is_dir() else PROJECT_ROOT


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="autotest", description="batch(.exe) 自動テスト実行ツール")
    parser.add_argument("command",
                        choices=["run", "validate", "list", "dbcheck", "new", "copy",
                                 "manual", "report", "gui"],
                        help="実行するコマンド（gui: 操作画面を開く / new: ひな形からケース作成 / "
                             "copy: 既存ケースの複製 / manual: 手動実施ケースの証跡採取 / "
                             "report: 証跡の結合 / dbcheck: SQL Server へ実際に接続できるか確認）")
    parser.add_argument("--id", default=None, help="new / copy で作るケースID")
    parser.add_argument("--template", default="normal",
                        help="new で使うひな形（templates/ の名前。既定: normal）")
    parser.add_argument("--from", dest="copy_from", default=None,
                        help="copy の複製元ケースID")
    parser.add_argument("--phase", choices=["before", "after"], default=None,
                        help="manual の採取段階（before: 実行前 / after: 実行後）")
    parser.add_argument("--session", default=None,
                        help="manual --phase after で使う作業フォルダ名（省略時は自動選択）")
    parser.add_argument("--force-new", action="store_true",
                        help="manual --phase before で、未完了の作業があっても新しく始める")
    parser.add_argument("--run", action="append", dest="runs",
                        help="report で結合する実行フォルダ（output 配下の名前。複数指定可）")
    parser.add_argument("--config", default=None,
                        help="設定ファイル（既定: config/settings.local.yaml があればそれ、無ければ settings.yaml）")
    parser.add_argument("--cases-dir", default=None, help="ケース定義フォルダ（既定: <プロジェクトルート>/cases）")
    parser.add_argument("--case", action="append", dest="cases", help="実行するケースID（複数指定可）")
    parser.add_argument("--tag", action="append", dest="tags", help="実行するタグ（複数指定可）")
    parser.add_argument("--out", default=None, help="Excel 出力先（既定: settings.excel.output_dir）")
    parser.add_argument("--offline", action="store_true", help="SQL Server の代わりに fixtures/ の CSV を使う")
    parser.add_argument("--dry-run", action="store_true",
                        help=".exe も DB 接続も行わず、流れと設定だけを確認する")
    parser.add_argument("--date", default=None,
                        help="パス・引数の {date} を展開する基準日（YYYYMMDD。既定は本日）")
    parser.add_argument("--keep-env", action="store_true",
                        help="teardown と設定ファイルの復元を行わない（失敗の調査用）")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="各工程の開始を逐次表示する（処理が止まる箇所の切り分け用）")
    # 注: NG が出ても後続ケースは常に実行する（全 NG を一度に把握するため）。
    #     以前あった --keep-going は「定義だけあって未使用」だったため削除した。
    return parser


def _display_width(text: str) -> int:
    """表示上の桁数。日本語などの全角文字を 2 桁として数える。

    str の長さで揃えると、全角を含む列がずれて読めなくなる。
    """
    import unicodedata

    return sum(2 if unicodedata.east_asian_width(ch) in "FWA" else 1 for ch in text)


def _pad(text: str, width: int) -> str:
    """表示幅を揃えて右側を空白で埋める。溢れる場合は 1 桁だけ空ける。"""
    return text + " " * max(1, width - _display_width(text))


def _list_cases(cases) -> int:
    """ケース一覧とタグ一覧を表示する。"""
    id_w = max([_display_width(c.case_id) for c in cases] + [10]) + 2
    tag_w = max([_display_width("/".join(c.tags)) for c in cases] + [8]) + 2

    print(_pad("ケースID", id_w) + _pad("区分", tag_w) + _pad("実行方式", 10) + "ケース名")
    print("-" * (id_w + tag_w + 50))
    for c in cases:
        mode = "手動" if c.is_manual else "自動"
        print(_pad(c.case_id, id_w) + _pad("/".join(c.tags), tag_w) + _pad(mode, 10) + c.name)

    # 使えるタグを提示する。--tag に何を渡せるか分からないと使えないため
    counts = {}
    for c in cases:
        for tag in c.tags:
            counts[tag] = counts.get(tag, 0) + 1
    if counts:
        print()
        print("タグ一覧（--tag で指定。複数指定するといずれかに一致するケースが対象）:")
        for tag in sorted(counts):
            print("  %s  (%d 件)" % (_pad(tag, 14), counts[tag]))
    print()
    print("%d 件" % len(cases))
    return 0


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
    # 「繋がったか」だけでは足りない。繋がった先が設定と違えば、別の DB を
    # 見た結果が正常な証跡として残る —— 最も気づきにくい誤りなので、
    # 成功扱いにはしない（終了コードでも区別できるようにする）。
    mismatch = False
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
            mismatch = True
            print("\n[警告] 接続先 DB が設定値 '%s' と異なります。" % db.get("database"))
            print("       このまま試験すると、別の DB を見た結果が証跡として残ります。")
            print("       settings の database.database と、ログインの既定データベースを")
            print("       確認してください。")
        cur.close()
    except Exception as exc:
        print("\n[警告] 接続はできましたが情報取得に失敗しました: %s" % exc)
    finally:
        conn.close()

    print("\n" + "=" * 66)
    if mismatch:
        print(" 接続はできましたが、接続先が設定と違います。上の警告を解消してください。")
        print("=" * 66)
        return 1
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

    print(f"設定ファイル : {args.config or _default_config()}")
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


def _scaffold(args, project_root: Path, cases_dir: Path) -> int:
    """new / copy コマンド。ケース定義の読み込みより先に処理する。

    new は「まだ存在しないケース」を作るコマンドなので、先に load_cases を
    通すと無関係な既存ケースの不備で作成できなくなる。
    """
    from . import scaffold

    if not args.id:
        print("[設定エラー] --id でケースIDを指定してください。\n"
              "  例: python -m autotest %s --id TC010_new%s"
              % (args.command, " --from TC001_normal" if args.command == "copy" else ""),
              file=sys.stderr)
        return 2

    config_hint = args.config or ""
    try:
        if args.command == "new":
            created = scaffold.create_case(project_root, cases_dir, args.id, args.template)
            print("ひな形 '%s' からケースを作成しました。" % args.template)
        else:
            if not args.copy_from:
                print("[設定エラー] --from で複製元のケースIDを指定してください。\n"
                      "  例: python -m autotest copy --from TC001_normal --id TC010_new",
                      file=sys.stderr)
                return 2
            created = scaffold.copy_case(project_root, cases_dir, args.copy_from, args.id)
            print("ケース %s を %s として複製しました。" % (args.copy_from, args.id))
    except ConfigError as exc:
        print("[設定エラー] %s" % exc, file=sys.stderr)
        return 2

    for path in created:
        print("  作成: %s" % path)
    print("\n次にやること:")
    for line in scaffold.next_steps(project_root, cases_dir, args.id, config_hint):
        print("  %s" % line)
    return 0


def _open_log(run_dir: Path, verbose: bool):
    """工程ログを開き、(書き込み関数, ファイル) を返す。

    処理が止まったとき「どこで止まっているか」を後から追えるようにする。
    手動実施は 2 回に分けて実行するので追記モードで開く。
    """
    run_dir.mkdir(parents=True, exist_ok=True)
    log_file = (run_dir / "run.log").open("a", encoding="utf-8")

    def write_log(message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        log_file.write("%s  %s\n" % (stamp, message))
        log_file.flush()  # ハング時に読めないと意味がないので都度フラッシュ

    return write_log, log_file


def _manual(args, settings, cases, out_dir: Path) -> int:
    """手動実施ケースの証跡採取（before / after）。"""
    from . import manual as manual_mod
    from . import result_store

    if not args.phase:
        print("[設定エラー] --phase before か --phase after を指定してください。", file=sys.stderr)
        return 2
    if not args.cases or len(args.cases) != 1:
        print("[設定エラー] --case でケースIDを 1 件だけ指定してください。", file=sys.stderr)
        return 2

    case_id = args.cases[0]
    target = [c for c in cases if c.case_id == case_id]
    if not target:
        print("[設定エラー] ケース %s が見つかりません（enabled: false ではありませんか）。" % case_id,
              file=sys.stderr)
        return 2
    case = target[0]
    if not case.is_manual:
        print("[設定エラー] %s は mode: manual ではありません。\n"
              "  自動実行してください: python -m autotest run --case %s" % (case_id, case_id),
              file=sys.stderr)
        return 2

    open_sessions = manual_mod.find_sessions(out_dir, case_id)

    if args.phase == "before":
        if open_sessions and not args.force_new:
            print("[設定エラー] このケースには未完了の手動作業があります:", file=sys.stderr)
            for path in open_sessions:
                print("    %s" % path, file=sys.stderr)
            print("  続きを採取する : python -m autotest manual --case %s --phase after" % case_id,
                  file=sys.stderr)
            print("  やり直す       : 上のコマンドに --force-new を付けて before から始める",
                  file=sys.stderr)
            return 2

        started = datetime.now()
        session_dir = out_dir / manual_mod.new_session_id(case_id, started)
        write_log, log_file = _open_log(session_dir, args.verbose)
        write_log("手動実施 before 開始 case=%s config=%s offline=%s"
                  % (case_id, settings.source, args.offline))
        print("=== 手動実施：実行前の採取  %s ===" % case_id)
        try:
            runner = CaseRunner(settings, session_dir, offline=args.offline,
                                default_date=args.date)
            runner._progress = _progress_printer(write_log, args.verbose, case_id)
            session = runner.run_manual_before(case, session_dir)
            session.config_path = str(settings.source)
            session.save(session_dir)
        except ConfigError as exc:
            write_log("失敗: %s" % exc)
            print("[設定エラー] %s" % exc, file=sys.stderr)
            return 2
        except Exception as exc:  # noqa: BLE001 - 利用者に原因を見せて終わる
            write_log("失敗: %s" % exc)
            print("[エラー] 実行前の採取に失敗しました: %s" % exc, file=sys.stderr)
            return 1
        finally:
            log_file.close()

        print("\n実行前の証跡を採取しました。")
        print("  作業フォルダ : %s" % session_dir)
        print("\n次にやること:")
        print("  1) batch を手で実行してください（必要な手動操作もこの間に）")
        print("  2) 終わったら実行後の採取:")
        print("       python -m autotest manual --case %s --phase after%s"
              % (case_id, _config_option(args)))
        return 0

    # --- after --------------------------------------------------------------
    if args.session:
        session_dir = out_dir / args.session
        if not session_dir.is_dir():
            print("[設定エラー] 作業フォルダがありません: %s" % session_dir, file=sys.stderr)
            return 2
    elif not open_sessions:
        print("[設定エラー] %s の未完了の手動作業がありません。\n"
              "  先に実行前の採取を行ってください:\n"
              "    python -m autotest manual --case %s --phase before%s"
              % (case_id, case_id, _config_option(args)), file=sys.stderr)
        return 2
    elif len(open_sessions) > 1:
        print("[設定エラー] 未完了の手動作業が複数あります。--session でどれか 1 つを指定してください:",
              file=sys.stderr)
        for path in open_sessions:
            print("    --session %s" % path.name, file=sys.stderr)
        return 2
    else:
        session_dir = open_sessions[0]

    write_log, log_file = _open_log(session_dir, args.verbose)
    write_log("手動実施 after 開始 case=%s session=%s" % (case_id, session_dir.name))
    print("=== 手動実施：実行後の採取  %s ===" % case_id)
    try:
        session = manual_mod.ManualSession.load(session_dir)
        if session.case_id != case_id:
            raise ConfigError(
                "作業フォルダのケース（%s）と指定（%s）が違います。" % (session.case_id, case_id))
        if session.phase != manual_mod.PHASE_BEFORE_DONE:
            raise ConfigError(
                "この作業フォルダは既に採取済みです（phase=%s）: %s\n"
                "  やり直す場合は before から始めてください。" % (session.phase, session_dir))

        runner = CaseRunner(settings, session_dir, offline=session.offline,
                            default_date=args.date)
        runner._progress = _progress_printer(write_log, args.verbose, case_id)
        result = runner.run_manual_after(case, session, session_dir)
    except ConfigError as exc:
        write_log("失敗: %s" % exc)
        print("[設定エラー] %s" % exc, file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        write_log("失敗: %s" % exc)
        print("[エラー] 実行後の採取に失敗しました: %s" % exc, file=sys.stderr)
        return 1
    finally:
        log_file.close()

    run = RunResult(
        run_id=session_dir.name,
        started_at=session.before_started_at or datetime.now(),
        finished_at=datetime.now(),
        env_name=str(settings.env.get("name", "")),
        tester=settings.tester,
        exe_path=str(settings.batch.get("exe_path", "")),
        db_server=str(settings.database.get("server", "")),
        db_name=str(settings.database.get("database", "")),
        filter_description="手動実施: %s" % case_id,
        total_available=1,
    )
    run.cases.append(result)

    result_store.save_case_result(result, session_dir)
    result_store.save_run_meta(run, session_dir)
    session.save(session_dir)   # phase = finished

    excel_path = build_workbook(run, out_dir / ("TestEvidence_%s.xlsx" % session_dir.name),
                                settings.excel)
    print("=" * 60)
    print("判定 : %s（%d 件が確認待ち / NG %d 件）"
          % (result.verdict, result.review_count, result.ng_count))
    print("証跡Excel: %s" % excel_path)
    print("作業フォルダ: %s" % session_dir)
    print("\n※ 手動実施ケースは自動では合格になりません。")
    print("  Excel の黄色い欄に確認結果を記入してください。")
    print("\n自動実行分と 1 冊にまとめる:")
    print("  python -m autotest report --run <自動実行のrun_id> --run %s --out output/提出用.xlsx"
          % session_dir.name)
    if result.ng_count:
        return 1
    return 3


def _progress_printer(write_log, verbose: bool, case_id: str):
    def on_step(step: str) -> None:
        write_log("    %s: %s" % (case_id, step))
        if verbose:
            print("      > %s" % step, flush=True)
        else:
            print("  %s ..." % step, flush=True)
    return on_step


def _config_option(args) -> str:
    return (" --config %s" % args.config) if args.config else ""


def _report(args, settings, out_dir: Path) -> int:
    """複数の実行結果を 1 冊の証跡 Excel にまとめる。"""
    from . import result_store

    if not args.runs:
        print("[設定エラー] --run で結合する実行フォルダを指定してください（複数指定可）。\n"
              "  例: python -m autotest report --run 20260807_101500_123 "
              "--run manual_TC008_20260807_103000_456 --out output/提出用.xlsx",
              file=sys.stderr)
        return 2

    runs = []
    sources = []
    try:
        for name in args.runs:
            run_dir = Path(name)
            if not run_dir.is_absolute():
                run_dir = out_dir / name
            if not run_dir.is_dir():
                raise ConfigError("実行フォルダがありません: %s" % run_dir)
            runs.append(result_store.load_run_dir(run_dir))
            sources.append(run_dir.name)
        merged = result_store.merge_runs(runs, sources)
    except ConfigError as exc:
        print("[設定エラー] %s" % exc, file=sys.stderr)
        return 2

    out_path = Path(args.out) if args.out else out_dir / ("TestEvidence_%s.xlsx" % merged.run_id)
    excel_path = build_workbook(merged, out_path, settings.excel)

    print("=" * 60)
    print("結合元 : %s" % ", ".join(sources))
    print("総合判定 : %s   OK %d / NG %d / 要確認 %d"
          % (merged.verdict, merged.ok_count, merged.ng_count, merged.review_count))
    print("証跡Excel: %s" % excel_path)
    if merged.verdict == REVIEW:
        return 3
    return 0 if merged.verdict == OK else 1


def main(argv: Optional[List[str]] = None) -> int:
    _prepare_console()
    args = build_parser().parse_args(argv)

    config_path = Path(args.config) if args.config else _default_config()
    project_root = _project_root_for(config_path)
    cases_dir = Path(args.cases_dir) if args.cases_dir else project_root / "cases"

    if args.command in ("new", "copy"):
        return _scaffold(args, project_root, cases_dir)

    if args.command == "gui":
        # 遅延 import。画面の無い環境（タスクスケジューラ等）で tkinter の
        # import に失敗しても、他のコマンドを巻き添えにしないため
        from . import gui
        return gui.main(project_root,
                        config_path=Path(args.config) if args.config else None,
                        cases_dir=cases_dir)

    try:
        settings = load_settings(config_path, project_root=project_root)
        settings.set_base_date(args.date)   # {date} の既定基準日
        cases = load_cases(cases_dir, only=args.cases, tags=args.tags)
        # 絞り込み前の全件数。証跡に「何件を実行しなかったか」を残すために使う
        total_available = len(load_cases(cases_dir)) if (args.cases or args.tags) else len(cases)
    except ConfigError as exc:
        print(f"[設定エラー] {exc}", file=sys.stderr)
        return 2

    if args.command == "list":
        return _list_cases(cases)

    if args.command == "dbcheck":
        return _dbcheck(settings, timeout_sec=int(settings.database.get("login_timeout_sec", 15)))

    if args.command == "validate":
        return _validate(settings, cases, args, project_root)

    # 成果物の既定の置き場。run では --out で差し替えられるが、report の --out は
    # 「出力する Excel のパス」なので、結合元を探す場所は常にこちらを使う
    default_out_dir = settings.resolve_dir(str(settings.excel.get("output_dir", "./output")))

    if args.command == "manual":
        return _manual(args, settings, cases,
                       Path(args.out).resolve() if args.out else default_out_dir)

    if args.command == "report":
        return _report(args, settings, default_out_dir)

    out_dir = Path(args.out).resolve() if args.out else default_out_dir

    # --- run --------------------------------------------------------------
    # mode: manual のケースは自動実行しない。ただし「実行しなかった」ことは
    # 証跡に残す —— 自動分が全 OK でも、それは全体の合格ではない
    manual_cases = [c for c in cases if c.is_manual]
    cases = [c for c in cases if not c.is_manual]
    if not cases:
        print("[設定エラー] 実行対象が手動実施ケースだけです: %s\n"
              "  手動ケースは次のコマンドで採取してください:\n"
              "    python -m autotest manual --case %s --phase before%s"
              % ([c.case_id for c in manual_cases],
                 manual_cases[0].case_id if manual_cases else "<ID>", _config_option(args)),
              file=sys.stderr)
        return 2
    if args.cases:
        # 手動ケースを名指しされた場合は、黙って飛ばさず案内する
        named = [c.case_id for c in manual_cases if c.case_id in set(args.cases)]
        if named:
            print("[設定エラー] %s は mode: manual です。run では実行できません。\n"
                  "  python -m autotest manual --case %s --phase before%s"
                  % (named, named[0], _config_option(args)), file=sys.stderr)
            return 2

    started = datetime.now()
    # ミリ秒まで含める。秒精度だと同一秒に 2 回起動した場合に
    # 証跡フォルダと Excel が同名になり、互いに上書きされる
    run_id = started.strftime("%Y%m%d_%H%M%S") + "_%03d" % (started.microsecond // 1000)
    run_dir = out_dir / run_id

    run = RunResult(
        run_id=run_id,
        started_at=started,
        env_name=str(settings.env.get("name", "")),
        tester=settings.tester,
        exe_path=str(settings.batch.get("exe_path", "")),
        db_server=str(settings.database.get("server", "")),
        db_name=str(settings.database.get("database", "")),
        filter_description=_filter_description(args),
        total_available=total_available,
    )
    run.manual_pending = [c.case_id for c in manual_cases]

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
        runner = CaseRunner(settings, run_dir, offline=args.offline,
                            dry_run=args.dry_run, default_date=args.date,
                            keep_env=args.keep_env)
        if args.keep_env:
            print("※ --keep-env: teardown と設定復元を行いません。調査後は手動で戻してください。")
            write_log("--keep-env 指定: teardown / 設定復元をスキップ")
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
            # ケースごとに即時保存する。途中で落ちてもここまでの結果は残り、
            # 後から report で 1 冊にまとめられる
            result_store.save_case_result(result, run_dir)

            detail = f" ({result.ng_count} 件 NG)" if result.ng_count else ""
            print(f"{'      => ' if args.verbose else ''}{result.verdict}{detail}")
            write_log(f"--- {case.case_id} 終了: {result.verdict}{detail}")
            if result.fatal_error:
                print(f"    実行時エラー: {result.fatal_error.splitlines()[0]}")
                write_log(f"    実行時エラー: {result.fatal_error}")
    finally:
        log_file.close()

    run.finished_at = datetime.now()
    result_store.save_run_meta(run, run_dir)

    # file_name_format では {run_id} のほか {filter} も使える
    # （例: "TestEvidence_{run_id}_{filter}.xlsx" -> TestEvidence_..._tag-受注.xlsx）
    file_name = str(settings.excel.get("file_name_format", "TestEvidence_{run_id}.xlsx")).format(
        run_id=run_id, filter=_filter_slug(args))
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

    print(f"総合判定 : {run.verdict}   OK {run.ok_count} / NG {run.ng_count} / 要確認 {run.review_count}")
    if run.manual_pending:
        # 自動分だけの結果を「全部通った」と読ませない
        print(f"※ 手動実施ケース {len(run.manual_pending)} 件は今回実行していません: "
              f"{', '.join(run.manual_pending)}")
        print(f"   python -m autotest manual --case {run.manual_pending[0]} --phase before"
              f"{_config_option(args)}")
    if run.verdict == "SKIP":
        print("※ 実際に判定されたケースがありません（全件 SKIP）。終了コードは 1 になります。")
    if run.review_count:
        print(f"※ {run.review_count} 件が人の確認待ちです。証跡 Excel の黄色い欄に")
        print("   確認結果を記入してください。確認が済むまで合格とはしません（終了コード 3）。")
    print(f"証跡Excel: {excel_path}")
    print(f"証跡画像 : {run_dir / 'evidence'}")
    print(f"回収物   : {run_dir / 'artifacts'}")
    # 要確認は「まだ確定していない」状態。NG（1）とも成功（0）とも区別できるよう
    # 専用の終了コードにする。CI 側で「人の確認待ち」を検知できる
    if run.verdict == REVIEW:
        return 3
    return 0 if run.verdict == OK else 1


if __name__ == "__main__":
    sys.exit(main())
