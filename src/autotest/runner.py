"""被テスト batch の起動と結果取得。

起動方式は 2 つ（settings の batch.type）:
  exe  … exe_path をそのまま起動する（C# 製 .exe など。既定）
  java … java [jvm_args] -jar <jar> [common_args] [execute.args] で起動する（Spring Boot 等）
コマンドの組み立ては build_command に一本化し、実行・preflight・証跡表示が同じ規則を使う。
"""

import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from .config import ConfigError, Settings, batch_program_key, batch_type
from .models import ExecutionInfo


def _resolve_exe(raw: str, project_root: Path) -> Path:
    """exe_path を解決する。

    - `{python}` は実行中の Python インタプリタに置換する。検証用の
      demo/fake_batch.py を OS を問わず起動できるようにするため。
    - 相対パスはプロジェクトルート基準にする（起動フォルダに依存させない）。
    """
    resolved = raw.replace("{python}", sys.executable)
    path = Path(resolved)
    if not path.is_absolute():
        path = (project_root / path).resolve()
    return path


def _resolve_java(batch: Dict[str, Any], project_root: Path) -> Optional[Path]:
    """type: java で起動する java 実行ファイルを決める。見つからなければ None。

    batch.java → batch.env_vars.JAVA_HOME → 環境変数 JAVA_HOME → PATH の順。
    PATH 任せにすると、端末ごとに違う JDK で試験した証跡ができてしまうため、
    実環境では batch.java か JAVA_HOME での明示を推奨する（validate が実際の解決先を表示する）。
    """
    raw = batch.get("java")
    if raw:
        return _resolve_exe(str(raw), project_root)
    exe_name = "java.exe" if os.name == "nt" else "java"
    java_home = (batch.get("env_vars") or {}).get("JAVA_HOME") or os.environ.get("JAVA_HOME")
    if java_home:
        return Path(str(java_home)) / "bin" / exe_name
    found = shutil.which("java")
    return Path(found) if found else None


def build_command(settings: Settings, batch: Dict[str, Any], args: List[str]) -> Tuple[Path, List[str]]:
    """(起動するプログラム, 引数リスト) を組み立てる。{date} 系はここで展開する。

    java の jar は絶対パスにして渡す。working_dir が jar の場所と違う構成でも、
    相対パスのまま渡すと起動フォルダ基準で解決されて別物（または無し）を指すため。
    """
    def expand(values) -> List[str]:
        return [settings.expand(str(a)) for a in (values or [])]

    if batch_type(batch) == "java":
        program = _resolve_java(batch, settings.project_root)
        if program is None:
            raise ConfigError(
                "java が見つかりません。settings の batch.java に java(.exe) のパスを書くか、"
                "環境変数 JAVA_HOME を設定してください。")
        jar = _resolve_exe(str(batch.get("jar") or ""), settings.project_root)
        prefix = expand(batch.get("jvm_args")) + ["-jar", str(jar)]
    else:
        program = _resolve_exe(str(batch.get("exe_path") or ""), settings.project_root)
        prefix = []
    return program, prefix + expand(batch.get("common_args")) + expand(args)


def default_working_dir(settings: Settings, batch: Dict[str, Any]) -> Path:
    """working_dir の既定。exe はその置き場所、java は jar の置き場所。"""
    raw = batch.get("working_dir")
    if raw:
        wd = Path(str(raw))
        return wd if wd.is_absolute() else (settings.project_root / wd).resolve()
    key = batch_program_key(batch)
    return _resolve_exe(str(batch.get(key) or ""), settings.project_root).parent


def describe_batch(settings: Settings, batch: Dict[str, Any]) -> str:
    """証跡のサマリに出す「何を起動したか」。java は jar を主にする。"""
    if batch_type(batch) == "java":
        jar = _resolve_exe(str(batch.get("jar") or ""), settings.project_root)
        return "%s（java -jar）" % jar
    return str(batch.get("exe_path", ""))


def batch_preflight(settings: Settings, batch: Dict[str, Any], label: str) -> List[str]:
    """起動に必要なもの（exe / java / jar / working_dir）が揃っているかを静的に確認する。

    主 batch と前置 batch で同じ規則を使う。フォルダを空にして DB を戻した後で
    「jar が無い」と分かる事故を、破壊的操作の前に止めるため。
    """
    problems: List[str] = []
    key = batch_program_key(batch)
    if not batch.get(key):
        problems.append("%s: %s が未設定です%s" % (label, key, "（type: java）" if key == "jar" else ""))
    elif batch_type(batch) == "java":
        jar = _resolve_exe(str(batch["jar"]), settings.project_root)
        if not jar.is_file():
            problems.append("%s: jar が見つかりません: %s" % (label, jar))
        java = _resolve_java(batch, settings.project_root)
        if java is None:
            problems.append("%s: java が見つかりません（batch.java か JAVA_HOME を設定してください）" % label)
        elif not java.exists():
            problems.append("%s: java が見つかりません: %s" % (label, java))
    else:
        exe = _resolve_exe(str(batch[key]), settings.project_root)
        if not exe.exists():
            problems.append("%s: 実行ファイルが見つかりません: %s" % (label, exe))
    working_dir = batch.get("working_dir")
    if working_dir:
        wd = Path(str(working_dir))
        if not wd.is_absolute():
            wd = (settings.project_root / wd).resolve()
        if not wd.is_dir():
            problems.append("%s: working_dir が存在しません: %s" % (label, wd))
    return problems


def _kill_process_tree(proc: "subprocess.Popen") -> None:
    """プロセスツリーごと確実に終了させる。

    .exe が別プロセスを起動していると、親だけ kill しても孫が
    標準出力パイプを掴んだままになり、communicate() が返らずハングする。
    Windows では taskkill /T、それ以外ではプロセスグループへシグナルを送る。
    """
    if proc.poll() is not None:
        return
    try:
        if os.name == "nt":
            subprocess.run(  # noqa: S603,S607
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
            )
        else:
            proc.kill()
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


def run_batch(
    settings: Settings,
    args: List[str],
    dry_run: bool = False,
    batch_name: Optional[str] = None,
    on_progress=None,
) -> ExecutionInfo:
    """batch を同期実行し、終了コード・標準出力・所要時間を記録する。

    batch_name を渡すと settings.yaml の batches.<name> の定義で実行する。
    未指定なら既定の batch: を使う。ログ切り出しに使うため開始/終了時刻はここで確定させる。
    """
    batch = settings.batch_profile(batch_name)
    program_key = batch_program_key(batch)
    if not batch.get(program_key):
        raise ConfigError(
            "batch の %s が未設定です（batch 名: %s）" % (program_key, batch_name or "（既定）")
        )

    # 引数にも {date} を使えるようにする（例: --date 20260803）。
    exe_path, full_args = build_command(settings, batch, args)
    command = subprocess.list2cmdline([str(exe_path), *full_args])

    info = ExecutionInfo(command=command, started_at=datetime.now())

    if dry_run:
        info.finished_at = datetime.now()
        info.exit_code = 0
        info.stdout = "[dry-run] batch は実行していません"
        return info

    where = "batch%s" % ("es." + batch_name if batch_name else "")
    if not exe_path.exists():
        raise ConfigError(
            f"batch の実行ファイルが見つかりません: {exe_path}\n"
            f"  settings.yaml の {where}.{'java' if program_key == 'jar' else 'exe_path'} を確認してください。"
        )
    if program_key == "jar":
        jar = _resolve_exe(str(batch["jar"]), settings.project_root)
        if not jar.is_file():
            raise ConfigError(f"batch の jar が見つかりません: {jar}\n"
                              f"  settings.yaml の {where}.jar を確認してください。")

    working_dir = default_working_dir(settings, batch)
    encoding = batch.get("console_encoding", "cp932")
    timeout = int(batch.get("timeout_sec", 600))

    env = os.environ.copy()
    env.update({str(k): str(v) for k, v in (batch.get("env_vars") or {}).items()})

    heartbeat = max(1, int(batch.get("heartbeat_sec", 10)))

    proc = subprocess.Popen(  # noqa: S603  実行対象はテスト設定で明示された .exe
        [str(exe_path)] + list(full_args),
        cwd=str(working_dir),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        # ★ 標準入力を空にする。これを渡さないと .exe は親の stdin を継承し、
        #   「続行するには何かキーを押してください」等の入力待ちで永久に止まる。
        #   DEVNULL なら即 EOF になり、対話待ちの batch もそのまま進む。
        stdin=subprocess.DEVNULL,
        env=env,
    )

    try:
        # 一定間隔で生存確認しながら待つ。完全にブロックすると
        # 「動いているのか固まったのか」が外から分からないため。
        waited = 0.0
        while True:
            try:
                out, err = proc.communicate(timeout=heartbeat)
                break
            except subprocess.TimeoutExpired:
                waited += heartbeat
                if waited >= timeout:
                    raise
                if on_progress:
                    on_progress("batch実行中… %d 秒経過（上限 %d 秒）" % (waited, timeout))

        info.exit_code = proc.returncode
        info.stdout = out.decode(encoding, errors="replace")
        info.stderr = err.decode(encoding, errors="replace")
    except subprocess.TimeoutExpired as exc:
        # ★ 子プロセスだけ kill しても、孫プロセスがパイプを掴んだままだと
        #   communicate() が返らず結局ハングする。プロセスツリーごと落とす。
        _kill_process_tree(proc)
        try:
            out, err = proc.communicate(timeout=10)
        except Exception:
            out, err = b"", b""
        exc.stdout, exc.stderr = out, err
        info.timed_out = True
        info.exit_code = None
        info.stdout = (exc.stdout or b"").decode(encoding, errors="replace")
        info.stderr = (exc.stderr or b"").decode(encoding, errors="replace")
        info.stderr += f"\n[タイムアウト] {timeout} 秒以内に終了しませんでした。"
    finally:
        info.finished_at = datetime.now()

    return info
