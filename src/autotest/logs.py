"""ログ収集: batch 実行で「増えた分だけ」を切り出す。

過去のログを丸ごと証跡に貼ると読めないので、当該実行の行だけを抽出する。
切り出しは信頼できる順に 3 段構え:
  1) バイトオフセット差分 — 実行前にファイルサイズを控え、増分だけ読む（最も確実）
  2) タイムスタンプ絞り込み — 新規/ローテートされたファイルは行頭時刻で絞る
  3) 全文（末尾 N 行） — 上記が使えない場合の保険
"""

import hashlib
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from .config import Settings
from .fsops import find_files, read_text_file
from .models import Table

# 行の重要度判定。日本語ログの「エラー=0」「エラー件数=0」を異常扱いしないよう、
# レベルタグ付きの記述を優先して拾う。settings.log.error_patterns で上書き可能。
DEFAULT_ERROR_PATTERNS = [
    r"\[(?:ERROR|FATAL|SEVERE)\]",
    r"(?:^|\s)(?:ERROR|FATAL)\s*[:：]",
    r"\b(?:Exception|StackTrace|Traceback)\b",
    r"異常終了|致命的|処理失敗",
    r"エラー(?!\s*(?:件数)?\s*[=:：]\s*0(?!\d))",
]
DEFAULT_WARN_PATTERNS = [
    r"\[(?:WARN|WARNING)\]",
    r"(?:^|\s)(?:WARN|WARNING)\s*[:：]",
    r"警告",
]


class LineClassifier:
    """ログ行を error / warn / 通常 に分類する。Excel の行色分けと画像描画で共用する。"""

    def __init__(self, cfg: Optional[dict] = None):
        cfg = cfg or {}
        self._error = [re.compile(p, re.IGNORECASE) for p in cfg.get("error_patterns", DEFAULT_ERROR_PATTERNS)]
        self._warn = [re.compile(p, re.IGNORECASE) for p in cfg.get("warn_patterns", DEFAULT_WARN_PATTERNS)]

    def classify(self, line: str) -> str:
        if any(r.search(line) for r in self._error):
            return "error"
        if any(r.search(line) for r in self._warn):
            return "warn"
        return ""


class LogSlice:
    """1 ログファイルから切り出した内容。"""

    def __init__(
        self,
        path: Path,
        lines: Optional[List[str]] = None,
        method: str = "",  # offset | timestamp | full
        total_lines_before_trim: int = 0,
    ) -> None:
        self.path = path
        self.lines = lines if lines is not None else []
        self.method = method
        self.total_lines_before_trim = total_lines_before_trim

    @property
    def text(self) -> str:
        return "\n".join(self.lines)


def snapshot_offsets(settings: Settings, log_dir: Optional[Path] = None) -> Dict[Path, int]:
    """実行前のログファイルサイズを控える。

    log_dir を渡すとそのフォルダを見る。batch ごとにログ出力先が違う場合に使う。
    """
    log_dir = log_dir or settings.resolve_dir("log_dir")
    patterns = settings.log.get("patterns", ["*.log"])
    offsets: Dict[Path, Tuple[int, int, str]] = {}
    for pattern in patterns:
        for path in find_files(log_dir, pattern):
            try:
                size = path.stat().st_size
                offsets[path] = (size, ) + _head_signature(path, size)
            except OSError:
                offsets[path] = (0, 0, "")
    return offsets


def _head_signature(path: Path, size: int, limit: int = 4096) -> Tuple[int, str]:
    """ファイル先頭のハッシュを返す（長さ, ダイジェスト）。

    ローテートで「同名だが中身が別のファイル」に入れ替わったことを検出する。
    サイズ比較だけでは、新ファイルが旧ファイルより大きい場合に
    旧オフセットから読んでしまい、新ログの先頭を取りこぼす。
    """
    head_len = min(size, limit)
    if head_len <= 0:
        return 0, ""
    with path.open("rb") as f:
        head = f.read(head_len)
    return head_len, hashlib.sha1(head).hexdigest()


def collect(
    settings: Settings,
    offsets: Dict[Path, Tuple[int, int, str]],
    started_at: Optional[datetime],
    finished_at: Optional[datetime],
    log_dir: Optional[Path] = None,
) -> List[LogSlice]:
    """実行後のログから当該実行分を切り出す。

    log_dir は snapshot_offsets と同じものを渡すこと（別フォルダだと増分が取れない）。

    ここでは表示用の中略は一切行わない。キーワード判定（check_keywords）は
    全行に対して行う必要があるため、表示の絞り込みは to_table() 側の責務。
    """
    log_dir = log_dir or settings.resolve_dir("log_dir")
    cfg = settings.log
    patterns = cfg.get("patterns", ["*.log"])
    encoding = cfg.get("encoding", "utf-8")
    fallbacks = cfg.get("encoding_fallbacks", ["cp932"])

    slices: List[LogSlice] = []
    seen: Set[Path] = set()
    for pattern in patterns:
        for path in find_files(log_dir, pattern):
            if path in seen:
                continue
            seen.add(path)

            before = offsets.get(path)
            try:
                current_size = path.stat().st_size
            except OSError:
                continue

            # ローテート検出: 先頭バイトのハッシュが変わっていたら別ファイル扱い
            if before is not None:
                size_before, head_len, head_sig = before
                if head_len > 0 and (current_size < head_len
                                     or _head_signature(path, current_size, head_len)[1] != head_sig):
                    before = None  # 中身が入れ替わった → 先頭から読み直す

            if before is not None and current_size > before[0]:
                lines = _read_from_offset(path, before[0], encoding, fallbacks)
                method = "offset"
            elif before is not None and current_size <= before[0]:
                continue  # 増えていない = この実行とは無関係
            else:
                text = read_text_file(path, encoding, fallbacks)
                lines = text.splitlines()
                method = "full"
                if cfg.get("timestamp_regex") and started_at:
                    filtered = _filter_by_time(lines, cfg, started_at, finished_at)
                    if filtered:
                        lines, method = filtered, "timestamp"

            if not lines:
                continue

            slices.append(LogSlice(path=path, lines=lines, method=method,
                                   total_lines_before_trim=len(lines)))

    return slices


def _read_from_offset(path: Path, offset: int, encoding: str, fallbacks: List[str]) -> List[str]:
    with path.open("rb") as f:
        f.seek(offset)
        raw = f.read()
    for enc in [encoding, *fallbacks]:
        try:
            return raw.decode(enc).splitlines()
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode(encoding or "utf-8", errors="replace").splitlines()


def _filter_by_time(
    lines: List[str],
    cfg: dict,
    started_at: datetime,
    finished_at: Optional[datetime],
) -> List[str]:
    """行頭タイムスタンプで実行時間帯の行に絞る。時刻の無い行は直前行に追従させる。"""
    regex = re.compile(cfg["timestamp_regex"])
    ts_format = cfg.get("timestamp_format", "%Y-%m-%d %H:%M:%S")
    margin = timedelta(seconds=int(cfg.get("slice_margin_sec", 3)))
    lower = started_at - margin
    upper = (finished_at or datetime.now()) + margin

    result: List[str] = []
    in_window = False
    for line in lines:
        m = regex.search(line)
        if m:
            try:
                ts = datetime.strptime(m.group("ts").replace("/", "-").replace("T", " "), ts_format)
                in_window = lower <= ts <= upper
            except ValueError:
                pass  # 解析できない行は直前の状態を引き継ぐ
        if in_window:
            result.append(line)
    return result


def to_table(
    sl: LogSlice,
    classifier: LineClassifier,
    highlight_keywords: Optional[List[str]] = None,
    max_lines: int = 500,
) -> Table:
    """ログ切片を Excel のネイティブ表に変換する。

    画像ではなくセルで出すことで、Excel 上での検索・コピー・フィルタが効く。
    ERROR/WARN 行と、期待値キーワードを含む行はセル色で示す。

    中略はここ（表示専用）でだけ行う。キーワード判定は check_keywords が
    LogSlice の全行に対して行うため、表示の絞り込みは判定に影響しない。
    """
    method_label = {
        "offset": "実行による増分",
        "timestamp": "実行時間帯で抽出",
        "full": "全文",
    }.get(sl.method, sl.method)

    keywords = [k for k in (highlight_keywords or []) if k]

    display: List[Tuple[Optional[int], str]] = [(i + 1, ln) for i, ln in enumerate(sl.lines)]
    if len(display) > max_lines:
        head = max_lines // 2
        omitted = len(display) - max_lines
        display = (display[:head]
                   + [(None, f"…（表示上の中略 {omitted} 行。判定は全行に対して実施済み）…")]
                   + display[-(max_lines - head):])
    rows: List[List[object]] = []
    marks: Dict[Tuple[int, int], str] = {}

    for i, (line_no, line) in enumerate(display):
        rows.append([line_no if line_no is not None else "…", line.expandtabs(4)])
        if line_no is None:
            marks[(i, 0)] = "hit"
            marks[(i, 1)] = "hit"
            continue
        level = classifier.classify(line)
        if not level and any(k in line for k in keywords):
            level = "hit"
        if level:
            marks[(i, 0)] = level
            marks[(i, 1)] = level

    return Table(
        title=sl.path.name,
        columns=["行", "内容"],
        rows=rows,
        cell_marks=marks,
        note=f"{sl.path}  —  切り出し方法: {method_label} / {sl.total_lines_before_trim} 行",
        mono_columns=[1],
    )


def check_keywords(
    slices: List[LogSlice],
    must_contain: List[str],
    must_not_contain: List[str],
) -> Tuple[List[str], List[str]]:
    """(未検出の必須キーワード, 検出された禁止キーワード) を返す。"""
    combined = "\n".join(s.text for s in slices)
    missing = [kw for kw in must_contain if kw not in combined]
    forbidden = [kw for kw in must_not_contain if kw in combined]
    return missing, forbidden
