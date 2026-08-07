"""テスト実行中に受け渡しするデータ構造。

Excel 出力層（excel.py）はこのモジュールの型だけを知っていればよく、
DB / ファイル / ログのどこから来た情報かを意識しない。

Python 3.6（Anaconda 5.2）を対象にしているため dataclass は使わず、
明示的な __init__ を持つ通常のクラスで書いている。可変デフォルト値
（リスト・辞書）は None を受けて __init__ 内で生成する。
"""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

OK = "OK"
NG = "NG"
SKIP = "SKIP"
ERROR = "ERROR"
# 自動判定だけでは確定させず、人が証跡を見て最終判定する項目。
# 自動比較の結果があっても OK にはせず、確認待ちとして残す。
REVIEW = "要確認"


class Table:
    """Excel にネイティブのセルとして書き出す表。DB スナップショットやログに使う。"""

    def __init__(
        self,
        title: str,
        columns: List[str],
        rows: List[List[Any]],
        cell_marks: Optional[Dict[Tuple[int, int], str]] = None,
        note: str = "",
        truncated_from: Optional[int] = None,
        mono_columns: Optional[List[int]] = None,
        mask_columns: Optional[List[str]] = None,
    ) -> None:
        self.title = title
        self.columns = columns
        self.rows = rows
        # (行index, 列index) -> "diff" | "extra" | "error" | "warn" | "hit"。Excel 側で色分けする
        self.cell_marks = cell_marks if cell_marks is not None else {}
        self.note = note
        self.truncated_from = truncated_from  # 元の行数（打ち切った場合）
        # 等幅フォントで表示する列インデックス（ログ本文など）
        self.mono_columns = mono_columns if mono_columns is not None else []
        # 出力時にマスクする列名。比較は生値で行い、Excel へ書くときだけ伏せる。
        # 比較前にマスクすると、値が違っても同じに見えて偽 OK になる
        self.mask_columns = mask_columns if mask_columns is not None else []

    @property
    def row_count(self) -> int:
        return self.truncated_from if self.truncated_from is not None else len(self.rows)


class ImageEvidence:
    """Excel に貼り付ける証跡画像。"""

    def __init__(self, title: str, path: Path, caption: str = "") -> None:
        self.title = title
        self.path = path
        self.caption = caption


class CheckResult:
    """1 件の判定項目。ケースの総合判定はこれらの AND。"""

    def __init__(
        self,
        name: str,
        category: str,  # exit_code | db | file | log
        verdict: str,  # OK | NG | SKIP | ERROR
        detail: str = "",
        diff_table: Optional[Table] = None,
    ) -> None:
        self.name = name
        self.category = category
        self.verdict = verdict
        self.detail = detail
        # NG の根拠として Excel に展開する差分表
        self.diff_table = diff_table

    @property
    def is_ng(self) -> bool:
        return self.verdict in (NG, ERROR)

    @property
    def needs_review(self) -> bool:
        return self.verdict == REVIEW


class ExecutionInfo:
    """.exe の実行結果。"""

    def __init__(
        self,
        command: str = "",
        started_at: Optional[datetime] = None,
        finished_at: Optional[datetime] = None,
        exit_code: Optional[int] = None,
        stdout: str = "",
        stderr: str = "",
        timed_out: bool = False,
    ) -> None:
        self.command = command
        self.started_at = started_at
        self.finished_at = finished_at
        self.exit_code = exit_code
        self.stdout = stdout
        self.stderr = stderr
        self.timed_out = timed_out

    @property
    def elapsed_sec(self) -> float:
        if not self.started_at or not self.finished_at:
            return 0.0
        return (self.finished_at - self.started_at).total_seconds()


class CaseResult:
    """1 テストケースの全成果。Excel の 1 シートに対応する。"""

    def __init__(
        self,
        case_id: str,
        name: str,
        description: str = "",
        tags: Optional[List[str]] = None,
    ) -> None:
        self.case_id = case_id
        self.name = name
        self.description = description
        self.tags = tags if tags is not None else []
        self.execution = ExecutionInfo()
        self.checks = []  # type: List[CheckResult]
        # DB スナップショット: テーブル名 -> Table
        self.db_before = {}  # type: Dict[str, Table]
        self.db_after = {}  # type: Dict[str, Table]
        # ログはネイティブ表として出す（画像化しない）
        self.log_tables = []  # type: List[Table]
        self.images = []  # type: List[ImageEvidence]
        # 保全した入出力ファイルの一覧（コピー先の絶対パス）
        self.saved_artifacts = []  # type: List[Path]
        self.fatal_error = ""

    @property
    def verdict(self) -> str:
        if self.fatal_error:
            return ERROR
        if not self.checks:
            return SKIP
        if any(c.is_ng for c in self.checks):
            return NG
        # 人の確認待ちが 1 つでもあれば、ケース全体を確定させない
        if any(c.needs_review for c in self.checks):
            return REVIEW
        if all(c.verdict == SKIP for c in self.checks):
            return SKIP
        return OK

    @property
    def ng_count(self) -> int:
        return sum(1 for c in self.checks if c.is_ng)

    @property
    def review_count(self) -> int:
        return sum(1 for c in self.checks if c.needs_review)


class RunResult:
    """1 回の実行（複数ケース）全体。Excel の サマリ シートに対応する。"""

    def __init__(
        self,
        run_id: str,
        started_at: datetime,
        finished_at: Optional[datetime] = None,
        env_name: str = "",
        tester: str = "",
        exe_path: str = "",
        db_server: str = "",
        db_name: str = "",
        filter_description: str = "",
        total_available: Optional[int] = None,
    ) -> None:
        self.run_id = run_id
        self.started_at = started_at
        self.finished_at = finished_at
        self.env_name = env_name
        self.tester = tester
        self.exe_path = exe_path
        self.db_server = db_server
        self.db_name = db_name
        # 絞り込み条件と、絞り込み前の全ケース数。
        # 証跡には「何を実行したか」だけでなく「何を実行しなかったか」も
        # 残す必要がある。一部だけ実行した結果を全体の合格と誤読させないため。
        self.filter_description = filter_description
        self.total_available = total_available
        self.cases = []  # type: List[CaseResult]
        # mode: manual のため今回実行しなかったケース ID。
        # 「自動分は全 OK」を「全部通った」と読ませないために証跡へ残す
        self.manual_pending = []  # type: List[str]

    @property
    def ok_count(self) -> int:
        return sum(1 for c in self.cases if c.verdict == OK)

    @property
    def ng_count(self) -> int:
        return sum(1 for c in self.cases if c.verdict in (NG, ERROR))

    @property
    def verdict(self) -> str:
        """実行全体の判定。

        「1 件も実行していない」「全件 SKIP」を OK にすると、設定ミスで
        何もテストされていないのに CI が緑になる（偽グリーン）。
        OK は「少なくとも 1 件が実際に OK 判定で、NG が 0 件」のときだけ。
        """
        if self.ng_count:
            return NG
        if not self.cases:
            return SKIP
        # 確認待ちが残っている間は合格にしない（人が見て初めて確定する）
        if any(c.verdict == REVIEW for c in self.cases):
            return REVIEW
        if all(c.verdict == SKIP for c in self.cases):
            return SKIP
        return OK

    @property
    def review_count(self) -> int:
        return sum(1 for c in self.cases if c.verdict == REVIEW)
