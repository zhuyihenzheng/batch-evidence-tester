"""期待値との突合。OK/NG 判定と、Excel に貼る差分表を作る。

DB は「キー列でレコードを対応付けてセル単位で比較」、ファイルは「行単位で比較」。
差分表は 1 行 1 相違点の形にする（横に長い表より Excel 上で追いやすい）。
"""

import csv
import difflib
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from .models import NG, OK, CheckResult, Table

DIFF_COLUMNS = ["区分", "キー", "列名", "期待値", "実績値"]
FILE_DIFF_COLUMNS = ["行", "区分", "期待値", "実績値"]

MAX_DIFF_ROWS = 200


def _norm(value: object) -> str:
    """比較用の正規化。前後空白と改行コードの差は無視する。"""
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def compare_db_table(
    name: str,
    actual: Table,
    expected: Table,
    keys: List[str],
    ignore_columns: Optional[List[str]] = None,
) -> CheckResult:
    """DB スナップショットと期待値 CSV を突合する。"""
    ignore = set(ignore_columns or [])
    check_name = f"DB照合: {name}"

    # 打ち切られたデータで OK を出すと、上限より後ろの異常が永久に見逃される。
    # 判定不能として明示的に NG にする（偽 OK の防止を正しさより優先しない）。
    for side, table in (("実績", actual), ("期待値", expected)):
        if table.truncated_from is not None:
            return CheckResult(
                check_name, "db", NG,
                f"{side}側のデータが {table.truncated_from} 行から {len(table.rows)} 行に"
                f"打ち切られているため判定できません。SQL の WHERE で対象を絞ってください。",
            )

    if not expected.columns:
        return CheckResult(check_name, "db", NG, "期待値ファイルが空です")

    missing_cols = [c for c in expected.columns if c not in actual.columns]
    if missing_cols:
        return CheckResult(
            check_name, "db", NG,
            f"実績側に存在しない列があります: {missing_cols}（SQL の SELECT 列を確認してください）",
        )

    if not keys:
        return _compare_by_position(check_name, "db", actual, expected, ignore)

    missing_keys = [k for k in keys if k not in expected.columns]
    if missing_keys:
        return CheckResult(check_name, "db", NG, f"キー列が期待値に存在しません: {missing_keys}")

    compare_cols = [c for c in expected.columns if c not in ignore and c not in keys]
    # 差分表へ出す値だけ伏せる。比較そのものは生値で行うので検出漏れはしない
    masked = set(actual.mask_columns) | set(expected.mask_columns)

    def shown(col: str, value: object) -> object:
        return "***MASKED***" if col in masked and value not in ("", None, "(NULL)") else value

    exp_index, exp_dups = _index_by_key(expected, keys)
    act_index, act_dups = _index_by_key(actual, keys)

    # dict 索引は同一キーで後勝ちになるため、重複があると異常な方のレコードが
    # 静かに上書きされて OK になり得る。重複自体を NG として明示する。
    if act_dups or exp_dups:
        parts = []
        if act_dups:
            parts.append(f"実績側にキー重複: {sorted(set(act_dups))[:10]}")
        if exp_dups:
            parts.append(f"期待値側にキー重複: {sorted(set(exp_dups))[:10]}")
        return CheckResult(
            check_name, "db", NG,
            " / ".join(parts) + f"（キー {keys} が一意になるよう SQL か期待値を見直してください）",
        )

    diff_rows: List[List[str]] = []

    for key, exp_row in exp_index.items():
        key_text = " / ".join(key)
        act_row = act_index.get(key)
        if act_row is None:
            diff_rows.append(["期待のみ", key_text, "(レコード全体)", _row_digest(expected.columns, exp_row, masked=masked), ""])
            continue
        for col in compare_cols:
            e = _norm(exp_row.get(col))
            a = _norm(act_row.get(col))
            if e != a:
                diff_rows.append(["相違", key_text, col,
                                  shown(col, exp_row.get(col, "")), shown(col, act_row.get(col, ""))])

    for key, act_row in act_index.items():
        if key not in exp_index:
            diff_rows.append(["実績のみ", " / ".join(key), "(レコード全体)", "", _row_digest(actual.columns, act_row, masked=masked)])

    return _build_result(
        check_name, "db", diff_rows, DIFF_COLUMNS,
        ok_detail=f"{len(exp_index)} 件一致（比較列 {len(compare_cols)} / 除外列 {sorted(ignore) or 'なし'}）",
    )


def _compare_by_position(check_name: str, category: str, actual: Table, expected: Table, ignore: Set[str]) -> CheckResult:
    """キー列が指定されない場合は行順で比較する（ORDER BY 済みが前提）。"""
    compare_cols = [c for c in expected.columns if c not in ignore]
    diff_rows: List[List[str]] = []
    # 比較は生値、差分表への出力だけ伏せる
    masked = set(actual.mask_columns) | set(expected.mask_columns)

    def shown(col: str, value: object) -> object:
        return "***MASKED***" if col in masked and value not in ("", None, "(NULL)") else value

    for i in range(max(len(expected.rows), len(actual.rows))):
        exp_row = dict(zip(expected.columns, expected.rows[i])) if i < len(expected.rows) else None
        act_row = dict(zip(actual.columns, actual.rows[i])) if i < len(actual.rows) else None
        if exp_row is None:
            diff_rows.append(["実績のみ", f"{i + 1}行目", "(レコード全体)", "",
                              _row_digest(actual.columns, act_row or {}, masked=masked)])
        elif act_row is None:
            diff_rows.append(["期待のみ", f"{i + 1}行目", "(レコード全体)",
                              _row_digest(expected.columns, exp_row, masked=masked), ""])
        else:
            for col in compare_cols:
                if _norm(exp_row.get(col)) != _norm(act_row.get(col)):
                    diff_rows.append(["相違", f"{i + 1}行目", col,
                                      shown(col, exp_row.get(col, "")), shown(col, act_row.get(col, ""))])

    return _build_result(
        check_name, category, diff_rows, DIFF_COLUMNS,
        ok_detail=f"{len(expected.rows)} 行一致（行順比較）",
    )


def compare_text_file(
    name: str,
    actual_path: Optional[Path],
    expected_path: Path,
    actual_text: Optional[str] = None,
    ignore_line_patterns: Optional[List[str]] = None,
) -> CheckResult:
    """出力ファイルを期待値ファイルと行単位で突合する。"""
    check_name = f"ファイル照合: {name}"

    if not expected_path.exists():
        return CheckResult(check_name, "file", NG, f"期待値ファイルがありません: {expected_path}")
    if actual_text is None and (actual_path is None or not actual_path.exists()):
        return CheckResult(check_name, "file", NG, f"実績ファイルが出力されていません: {actual_path}")

    # actual_text 未指定なら actual_path から読む（文字コードは順に試す）
    if actual_text is None:
        from .fsops import read_text_file  # 循環 import 回避のため遅延 import

        actual_text = read_text_file(actual_path, "utf-8-sig", ["cp932", "utf-8"])

    expected_lines = expected_path.read_text(encoding="utf-8-sig").splitlines()
    actual_lines = actual_text.splitlines()

    if ignore_line_patterns:
        import re

        regexes = [re.compile(p) for p in ignore_line_patterns]
        keep = lambda ln: not any(r.search(ln) for r in regexes)  # noqa: E731
        expected_lines = [ln for ln in expected_lines if keep(ln)]
        actual_lines = [ln for ln in actual_lines if keep(ln)]

    diff_rows: List[List[str]] = []
    matcher = difflib.SequenceMatcher(None, [_norm(x) for x in expected_lines], [_norm(x) for x in actual_lines])
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "replace":
            for offset in range(max(i2 - i1, j2 - j1)):
                e = expected_lines[i1 + offset] if i1 + offset < i2 else ""
                a = actual_lines[j1 + offset] if j1 + offset < j2 else ""
                diff_rows.append([str(i1 + offset + 1), "相違", e, a])
        elif tag == "delete":
            for i in range(i1, i2):
                diff_rows.append([str(i + 1), "期待のみ", expected_lines[i], ""])
        elif tag == "insert":
            for j in range(j1, j2):
                diff_rows.append([str(j + 1), "実績のみ", "", actual_lines[j]])

    return _build_result(
        check_name, "file", diff_rows, FILE_DIFF_COLUMNS,
        ok_detail=f"{len(expected_lines)} 行一致",
    )


def compare_csv_file(name: str, actual_path: Path, expected_path: Path, keys: List[str], encoding: str = "utf-8") -> CheckResult:
    """CSV 出力をキー付きで突合する（列順の違いを吸収したい場合に使う）。"""
    if not actual_path.exists():
        return CheckResult(f"ファイル照合: {name}", "file", NG, f"実績ファイルが出力されていません: {actual_path}")
    actual = _read_csv_as_table(actual_path, encoding)
    expected = _read_csv_as_table(expected_path, "utf-8-sig")
    result = compare_db_table(name, actual, expected, keys)
    result.category = "file"
    result.name = f"ファイル照合: {name}"
    return result


def _read_csv_as_table(path: Path, encoding: str) -> Table:
    with path.open("r", encoding=encoding, newline="", errors="replace") as f:
        rows = list(csv.reader(f))
    if not rows:
        return Table(title=path.name, columns=[], rows=[])
    return Table(title=path.name, columns=rows[0], rows=[list(r) for r in rows[1:]])


def _index_by_key(
    table: Table, keys: List[str]
) -> Tuple[Dict[Tuple[str, ...], Dict[str, object]], List[str]]:
    """キーで索引化する。戻り値は (索引, 重複したキーの一覧)。

    dict は同一キーで後勝ちになるため、重複の有無を必ず呼び出し側へ返す。
    重複を握りつぶすと「異常なレコードが正常なレコードで上書きされて OK」
    という偽 OK が起きる。
    """
    key_pos = [table.columns.index(k) for k in keys if k in table.columns]
    index: Dict[Tuple[str, ...], Dict[str, object]] = {}
    duplicates: List[str] = []
    for row in table.rows:
        key = tuple(_norm(row[i]) for i in key_pos)
        if key in index:
            duplicates.append(" / ".join(key))
        index[key] = dict(zip(table.columns, row))
    return index, duplicates


def _row_digest(columns: List[str], row: Dict[str, object], limit: int = 6,
                masked: Optional[Set[str]] = None) -> str:
    masked = masked or set()
    parts = [f"{c}=" + ("***MASKED***" if c in masked else str(row.get(c, "")))
             for c in columns[:limit]]
    if len(columns) > limit:
        parts.append("...")
    return ", ".join(parts)


def _build_result(name: str, category: str, diff_rows: List[List[str]], columns: List[str], ok_detail: str) -> CheckResult:
    if not diff_rows:
        return CheckResult(name, category, OK, ok_detail)

    total = len(diff_rows)
    shown = diff_rows[:MAX_DIFF_ROWS]
    marks = {(r, c): "diff" for r in range(len(shown)) for c in range(len(columns))}
    table = Table(
        title=f"{name} 差分",
        columns=columns,
        rows=shown,
        cell_marks=marks,
        truncated_from=total if total > MAX_DIFF_ROWS else None,
    )
    detail = f"{total} 件の相違があります" + (f"（先頭 {MAX_DIFF_ROWS} 件を表示）" if total > MAX_DIFF_ROWS else "")
    return CheckResult(name, category, NG, detail, diff_table=table)
