"""判定結果の保存と復元。

用途は 2 つ:
  1. 自動実行と手動実施の結果を 1 冊の Excel にまとめる（autotest report）
  2. 実行が途中で落ちても、そこまでのケースの結果を残す

★ なぜ「既存の Excel に追記」ではないのか
  openpyxl 2.5（Anaconda 5.2 同梱）は、画像を含む xlsx を読み込んで保存し直すと
  画像を落とす。証跡画像はこのツールの成果物そのものなので、追記方式は
  「証跡を静かに破壊する」ことと同義になる。よって結合は必ず
  「保存しておいた結果から、新しいブックをゼロから組み直す」方式で行う。

★ 復元したデータで比較をやり直すことはしない
  判定は採取した時点で確定させる。後から作り直せる仕組みにすると、
  証跡の結論が実行後に変わり得ることになってしまう。ここでの復元は
  Excel を組み直すための表示専用。
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from .config import ConfigError
from .models import CaseResult, CheckResult, ExecutionInfo, ImageEvidence, RunResult, Table

FORMAT_VERSION = 1
RESULTS_DIR_NAME = "results"
RUN_META_NAME = "_run.json"

_TIME_FORMAT = "%Y-%m-%dT%H:%M:%S.%f"


def _dumps_time(value):
    # type: (Optional[datetime]) -> Optional[str]
    return value.strftime(_TIME_FORMAT) if value else None


def _loads_time(value):
    # type: (Optional[str]) -> Optional[datetime]
    return datetime.strptime(value, _TIME_FORMAT) if value else None


def _relative(path, base):
    # type: (Path, Path) -> str
    """base からの相対パスにする。フォルダごと移動しても読めるようにするため。

    証跡フォルダは転送されたり別ドライブへコピーされたりする。絶対パスで
    持つと移動先で画像が見つからず、貼り付けが黙って抜ける。
    """
    try:
        return str(Path(path).relative_to(base)).replace("\\", "/")
    except ValueError:
        # 実行フォルダの外を指している場合は絶対パスのまま残す
        return str(path)


def _resolve(rel, base):
    # type: (str, Path) -> Path
    path = Path(rel)
    return path if path.is_absolute() else (base / path)


# =============================================================================
# Table
# =============================================================================


def _table_to_dict(table):
    # type: (Table) -> Dict[str, Any]
    # cell_marks のキーは (行, 列) のタプル。JSON はタプルキーを持てないので
    # "行,列" の文字列にする
    marks = dict(("%d,%d" % (r, c), v) for (r, c), v in table.cell_marks.items())
    return {
        "title": table.title,
        "columns": list(table.columns),
        "rows": [list(row) for row in table.rows],
        "cell_marks": marks,
        "note": table.note,
        "truncated_from": table.truncated_from,
        "mono_columns": list(table.mono_columns),
        "mask_columns": list(table.mask_columns),
    }


def _table_from_dict(data):
    # type: (Dict[str, Any]) -> Table
    marks = {}
    for key, value in (data.get("cell_marks") or {}).items():
        row, col = key.split(",")
        marks[(int(row), int(col))] = value
    return Table(
        title=data.get("title", ""),
        columns=list(data.get("columns") or []),
        rows=[list(row) for row in (data.get("rows") or [])],
        cell_marks=marks,
        note=data.get("note", ""),
        truncated_from=data.get("truncated_from"),
        mono_columns=list(data.get("mono_columns") or []),
        mask_columns=list(data.get("mask_columns") or []),
    )


# =============================================================================
# CaseResult
# =============================================================================


def case_result_to_dict(result, run_dir):
    # type: (CaseResult, Path) -> Dict[str, Any]
    execution = result.execution
    return {
        "format_version": FORMAT_VERSION,
        "case_id": result.case_id,
        "name": result.name,
        "description": result.description,
        "tags": list(result.tags),
        "fatal_error": result.fatal_error,
        "execution": {
            "command": execution.command,
            "started_at": _dumps_time(execution.started_at),
            "finished_at": _dumps_time(execution.finished_at),
            "exit_code": execution.exit_code,
            "stdout": execution.stdout,
            "stderr": execution.stderr,
            "timed_out": execution.timed_out,
        },
        "checks": [
            {
                "name": c.name,
                "category": c.category,
                "verdict": c.verdict,
                "detail": c.detail,
                "diff_table": _table_to_dict(c.diff_table) if c.diff_table else None,
            }
            for c in result.checks
        ],
        "db_before": dict((k, _table_to_dict(v)) for k, v in result.db_before.items()),
        "db_after": dict((k, _table_to_dict(v)) for k, v in result.db_after.items()),
        "log_tables": [_table_to_dict(t) for t in result.log_tables],
        "images": [
            {"title": im.title, "path": _relative(im.path, run_dir), "caption": im.caption}
            for im in result.images
        ],
        "saved_artifacts": [_relative(p, run_dir) for p in result.saved_artifacts],
    }


def case_result_from_dict(data, run_dir):
    # type: (Dict[str, Any], Path) -> CaseResult
    version = data.get("format_version")
    if version != FORMAT_VERSION:
        raise ConfigError(
            "判定結果の書式バージョンが違います（ファイル: %r / 想定: %d）。\n"
            "  ツールを更新した場合、その実行結果は再実行して採り直してください。"
            % (version, FORMAT_VERSION))

    result = CaseResult(
        case_id=str(data.get("case_id") or ""),
        name=str(data.get("name") or ""),
        description=str(data.get("description") or ""),
        tags=list(data.get("tags") or []),
    )
    result.fatal_error = str(data.get("fatal_error") or "")

    ex = data.get("execution") or {}
    result.execution = ExecutionInfo(
        command=str(ex.get("command") or ""),
        started_at=_loads_time(ex.get("started_at")),
        finished_at=_loads_time(ex.get("finished_at")),
        exit_code=ex.get("exit_code"),
        stdout=str(ex.get("stdout") or ""),
        stderr=str(ex.get("stderr") or ""),
        timed_out=bool(ex.get("timed_out", False)),
    )

    for item in (data.get("checks") or []):
        diff = item.get("diff_table")
        result.checks.append(CheckResult(
            name=str(item.get("name") or ""),
            category=str(item.get("category") or ""),
            verdict=str(item.get("verdict") or ""),
            detail=str(item.get("detail") or ""),
            diff_table=_table_from_dict(diff) if diff else None,
        ))

    for key, target in (("db_before", result.db_before), ("db_after", result.db_after)):
        for name, table in (data.get(key) or {}).items():
            target[name] = _table_from_dict(table)

    result.log_tables = [_table_from_dict(t) for t in (data.get("log_tables") or [])]
    result.images = [
        ImageEvidence(title=str(im.get("title") or ""),
                      path=_resolve(str(im.get("path") or ""), run_dir),
                      caption=str(im.get("caption") or ""))
        for im in (data.get("images") or [])
    ]
    result.saved_artifacts = [_resolve(p, run_dir) for p in (data.get("saved_artifacts") or [])]
    return result


def save_case_result(result, run_dir):
    # type: (CaseResult, Path) -> Path
    """1 ケース分の結果を run_dir/results/<case_id>.json へ書く。"""
    results_dir = run_dir / RESULTS_DIR_NAME
    results_dir.mkdir(parents=True, exist_ok=True)
    path = results_dir / (result.case_id + ".json")
    with path.open("w", encoding="utf-8") as f:
        json.dump(case_result_to_dict(result, run_dir), f,
                  ensure_ascii=False, indent=1, sort_keys=True)
    return path


# =============================================================================
# RunResult（実行全体のメタ情報）
# =============================================================================


def save_run_meta(run, run_dir):
    # type: (RunResult, Path) -> Path
    results_dir = run_dir / RESULTS_DIR_NAME
    results_dir.mkdir(parents=True, exist_ok=True)
    path = results_dir / RUN_META_NAME
    data = {
        "format_version": FORMAT_VERSION,
        "run_id": run.run_id,
        "started_at": _dumps_time(run.started_at),
        "finished_at": _dumps_time(run.finished_at),
        "env_name": run.env_name,
        "tester": run.tester,
        "exe_path": run.exe_path,
        "db_server": run.db_server,
        "db_name": run.db_name,
        "filter_description": run.filter_description,
        "total_available": run.total_available,
        "manual_pending": list(run.manual_pending),
        "case_order": [c.case_id for c in run.cases],
    }
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)
    return path


def load_run_dir(run_dir):
    # type: (Path) -> RunResult
    """実行フォルダ（または手動作業フォルダ）から RunResult を復元する。"""
    results_dir = run_dir / RESULTS_DIR_NAME
    meta_path = results_dir / RUN_META_NAME
    if not meta_path.exists():
        raise ConfigError(
            "結果情報がありません: %s\n"
            "  証跡を結合できるのは、このツールで実行したフォルダ（output/<run_id>/ または\n"
            "  output/manual_<ケースID>_<日時>/）だけです。" % meta_path)
    try:
        with meta_path.open("r", encoding="utf-8") as f:
            meta = json.load(f)
    except ValueError as exc:
        raise ConfigError("%s を読めませんでした（壊れている可能性）: %s" % (meta_path, exc))

    if meta.get("format_version") != FORMAT_VERSION:
        raise ConfigError(
            "%s の書式バージョンが違います（ファイル: %r / 想定: %d）。"
            % (meta_path, meta.get("format_version"), FORMAT_VERSION))

    run = RunResult(
        run_id=str(meta.get("run_id") or run_dir.name),
        started_at=_loads_time(meta.get("started_at")) or datetime.now(),
        finished_at=_loads_time(meta.get("finished_at")),
        env_name=str(meta.get("env_name") or ""),
        tester=str(meta.get("tester") or ""),
        exe_path=str(meta.get("exe_path") or ""),
        db_server=str(meta.get("db_server") or ""),
        db_name=str(meta.get("db_name") or ""),
        filter_description=str(meta.get("filter_description") or ""),
        total_available=meta.get("total_available"),
    )
    run.manual_pending = list(meta.get("manual_pending") or [])

    # 実行順を保つ。証跡の並びが実行のたびに変わると差分が読みにくい
    order = list(meta.get("case_order") or [])
    files = dict((p.stem, p) for p in results_dir.glob("*.json") if p.name != RUN_META_NAME)
    for case_id in order:
        path = files.pop(case_id, None)
        if path is None:
            # 実行途中で落ちた場合はここに来る。黙って落とさず記録だけ残す
            continue
        run.cases.append(_load_case_file(path, run_dir))
    for case_id in sorted(files):
        run.cases.append(_load_case_file(files[case_id], run_dir))
    return run


def _load_case_file(path, run_dir):
    # type: (Path, Path) -> CaseResult
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except ValueError as exc:
        raise ConfigError("%s を読めませんでした（壊れている可能性）: %s" % (path, exc))
    return case_result_from_dict(data, run_dir)


def merge_runs(runs, sources):
    # type: (List[RunResult], List[str]) -> RunResult
    """複数の実行結果を 1 つにまとめる。

    ケース ID の重複は黙って上書きせずエラーにする。どちらの結果が
    採用されたか分からない証跡は、証跡として成立しない。
    """
    if not runs:
        raise ConfigError("結合する実行結果がありません。")

    head = runs[0]
    merged = RunResult(
        run_id="merged_" + datetime.now().strftime("%Y%m%d_%H%M%S"),
        started_at=min(r.started_at for r in runs),
        finished_at=max([r.finished_at for r in runs if r.finished_at] or [None]),
        env_name=head.env_name,
        tester=head.tester,
        exe_path=head.exe_path,
        db_server=head.db_server,
        db_name=head.db_name,
        filter_description="統合レポート: " + ", ".join(sources),
    )

    seen = {}
    for run, source in zip(runs, sources):
        for case in run.cases:
            if case.case_id in seen:
                raise ConfigError(
                    "ケース %s が複数の実行結果に含まれています（%s と %s）。\n"
                    "  どちらの結果を採用したか分からなくなるため、結合対象を絞ってください。"
                    % (case.case_id, seen[case.case_id], source))
            seen[case.case_id] = source
            merged.cases.append(case)
        for case_id in run.manual_pending:
            if case_id not in seen and case_id not in merged.manual_pending:
                merged.manual_pending.append(case_id)

    # 結合後に採取済みになった手動ケースは「未実施」から外す
    merged.manual_pending = [c for c in merged.manual_pending if c not in seen]
    merged.total_available = len(merged.cases) + len(merged.manual_pending)

    # 実行環境が混ざっている場合は証跡に明記する（別環境の結果を
    # 1 冊にまとめると、どの環境の結果か分からなくなる）
    envs = sorted(set(r.env_name for r in runs if r.env_name))
    if len(envs) > 1:
        merged.env_name = " / ".join(envs) + "（★複数環境の結果が混在しています）"
    return merged
