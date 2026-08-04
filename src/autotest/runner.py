"""被テスト batch (.exe) の起動と結果取得。"""

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from .config import ConfigError, Settings
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


def run_batch(
    settings: Settings,
    args: List[str],
    dry_run: bool = False,
    batch_name: Optional[str] = None,
) -> ExecutionInfo:
    """batch を同期実行し、終了コード・標準出力・所要時間を記録する。

    batch_name を渡すと settings.yaml の batches.<name> の定義で実行する。
    未指定なら既定の batch: を使う。ログ切り出しに使うため開始/終了時刻はここで確定させる。
    """
    batch = settings.batch_profile(batch_name)
    if not batch.get("exe_path"):
        raise ConfigError(
            "batch の exe_path が未設定です（batch 名: %s）" % (batch_name or "（既定）")
        )

    exe_path = _resolve_exe(str(batch["exe_path"]), settings.project_root)
    common_args = [str(a) for a in (batch.get("common_args") or [])]
    full_args = common_args + [str(a) for a in args]
    command = subprocess.list2cmdline([str(exe_path), *full_args])

    info = ExecutionInfo(command=command, started_at=datetime.now())

    if dry_run:
        info.finished_at = datetime.now()
        info.exit_code = 0
        info.stdout = "[dry-run] .exe は実行していません"
        return info

    if not exe_path.exists():
        raise ConfigError(
            f"batch の実行ファイルが見つかりません: {exe_path}\n"
            f"  settings.yaml の batch{'es.' + batch_name if batch_name else ''}.exe_path を確認してください。"
        )

    working_dir = Path(batch.get("working_dir") or exe_path.parent)
    if not working_dir.is_absolute():
        working_dir = (settings.project_root / working_dir).resolve()
    encoding = batch.get("console_encoding", "cp932")
    timeout = int(batch.get("timeout_sec", 600))

    env = os.environ.copy()
    env.update({str(k): str(v) for k, v in (batch.get("env_vars") or {}).items()})

    try:
        # capture_output= は Python 3.7 以降のため、3.6 でも動く書き方にしている
        proc = subprocess.run(  # noqa: S603  実行対象はテスト設定で明示された .exe
            [str(exe_path)] + list(full_args),
            cwd=str(working_dir),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
            env=env,
            check=False,
        )
        info.exit_code = proc.returncode
        info.stdout = proc.stdout.decode(encoding, errors="replace")
        info.stderr = proc.stderr.decode(encoding, errors="replace")
    except subprocess.TimeoutExpired as exc:
        info.timed_out = True
        info.exit_code = None
        info.stdout = (exc.stdout or b"").decode(encoding, errors="replace")
        info.stderr = (exc.stderr or b"").decode(encoding, errors="replace")
        info.stderr += f"\n[タイムアウト] {timeout} 秒以内に終了しませんでした。"
    finally:
        info.finished_at = datetime.now()

    return info
