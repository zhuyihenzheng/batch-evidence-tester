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
    parser.add_argument("command", choices=["run", "validate", "list"], help="実行するコマンド")
    parser.add_argument("--config", default=str(PROJECT_ROOT / "config" / "settings.yaml"), help="設定ファイル")
    parser.add_argument("--cases-dir", default=None, help="ケース定義フォルダ（既定: <プロジェクトルート>/cases）")
    parser.add_argument("--case", action="append", dest="cases", help="実行するケースID（複数指定可）")
    parser.add_argument("--tag", action="append", dest="tags", help="実行するタグ（複数指定可）")
    parser.add_argument("--out", default=None, help="Excel 出力先（既定: settings.excel.output_dir）")
    parser.add_argument("--offline", action="store_true", help="SQL Server の代わりに fixtures/ の CSV を使う")
    parser.add_argument("--dry-run", action="store_true", help=".exe を起動せず、フォルダ操作も行わない")
    # 注: NG が出ても後続ケースは常に実行する（全 NG を一度に把握するため）。
    #     以前あった --keep-going は「定義だけあって未使用」だったため削除した。
    return parser


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
    for alias, path in settings.path_aliases.items():
        mark = "存在" if path.is_dir() else "未作成（run 時に自動作成）"
        if not path.is_dir():
            warnings.append(f"paths.{alias} は未作成です: {path}")
        print(f"  paths.{alias:<16} {path}  ({mark})")

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

    runner = CaseRunner(settings, run_dir, offline=args.offline, dry_run=args.dry_run)
    for i, case in enumerate(cases, start=1):
        print(f"[{i}/{len(cases)}] {case.case_id} {case.name} ... ", end="", flush=True)
        result = runner.run(case)
        run.cases.append(result)
        detail = f" ({result.ng_count} 件 NG)" if result.ng_count else ""
        print(f"{result.verdict}{detail}")
        if result.fatal_error:
            print(f"    実行時エラー: {result.fatal_error.splitlines()[0]}")

    run.finished_at = datetime.now()

    file_name = str(settings.excel.get("file_name_format", "TestEvidence_{run_id}.xlsx")).format(run_id=run_id)
    excel_path = build_workbook(run, out_dir / file_name, settings.excel)

    print("=" * 60)
    print(f"総合判定 : {run.verdict}   OK {run.ok_count} / NG {run.ng_count}")
    if run.verdict == "SKIP":
        print("※ 実際に判定されたケースがありません（全件 SKIP）。終了コードは 1 になります。")
    print(f"証跡Excel: {excel_path}")
    print(f"証跡画像 : {run_dir / 'evidence'}")
    print(f"回収物   : {run_dir / 'artifacts'}")
    return 0 if run.verdict == OK else 1


if __name__ == "__main__":
    sys.exit(main())
