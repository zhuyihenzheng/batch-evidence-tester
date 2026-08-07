"""手動実施ケース（mode: manual）の中間状態。

自動では再現できないケース——実行中に人の操作が要る、タイミングを合わせる必要が
ある、といったもの——は 2 段階に分けて証跡を採る。

    autotest manual --case X --phase before   … 前処理と実行前スナップショット
    （人が手で batch を動かす。数分後でも、別のコンソールからでもよい）
    autotest manual --case X --phase after    … ログ増分・成果回収・実行後スナップショット

before と after は別プロセスなので、間の状態をディスクに残す必要がある。
このモジュールはその受け渡しだけを担当する（採取そのものは orchestrator）。

引き継ぐ状態は意図的に最小限にしてある:
  - ログの読み取り位置（バイトオフセット + 先頭署名）
  - 基準日（{date} の展開に使う。跨日で before と after がずれると
    RESULT_{date}.csv のようなパターンが静かに外れる）
  - 実行前スナップショットの保存先（CSV は別ファイルとして残す）
差し替えたファイルの退避情報や DB ロックは持たない。それらは「必ず元に戻す」
仕組みが自動実行の finally に依存しており、2 プロセスに跨がせると
「戻し損ねたまま」の経路ができてしまうため、preflight で使用自体を禁じている。
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from .config import ConfigError

# session.json の書式バージョン。読み込み側は不一致を黙って通さない
# （古い session を新しいコードで読むと、意味の変わったフィールドを
#  誤解したまま証跡を作ってしまう）
FORMAT_VERSION = 1

PHASE_BEFORE_DONE = "before_done"
PHASE_FINISHED = "finished"

_TIME_FORMAT = "%Y-%m-%dT%H:%M:%S.%f"


def _dumps_time(value: Optional[datetime]) -> Optional[str]:
    return value.strftime(_TIME_FORMAT) if value else None


def _loads_time(value: Optional[str]) -> Optional[datetime]:
    return datetime.strptime(value, _TIME_FORMAT) if value else None


class ManualSession(object):
    """手動ケースの before / after を繋ぐ状態。"""

    def __init__(
        self,
        session_id,          # type: str
        case_id,             # type: str
        case_source="",      # type: str
        config_path="",      # type: str
        offline=False,       # type: bool
        base_date="",        # type: str
        batch_name=None,     # type: Optional[str]
        log_dir="",          # type: str
        phase=PHASE_BEFORE_DONE,   # type: str
        before_started_at=None,    # type: Optional[datetime]
        before_finished_at=None,   # type: Optional[datetime]
        log_offsets=None,          # type: Optional[List[Dict[str, Any]]]
        snapshot_tables=None,      # type: Optional[List[str]]
        before_images=None,        # type: Optional[List[Dict[str, Any]]]
    ):
        self.session_id = session_id
        self.case_id = case_id
        self.case_source = case_source
        self.config_path = config_path
        self.offline = offline
        self.base_date = base_date
        self.batch_name = batch_name
        self.log_dir = log_dir
        self.phase = phase
        self.before_started_at = before_started_at
        self.before_finished_at = before_finished_at
        self.log_offsets = log_offsets if log_offsets is not None else []
        self.snapshot_tables = snapshot_tables if snapshot_tables is not None else []
        # 実行前に撮ったフォルダ一覧画像。画像そのものは session フォルダに
        # 残っているが、どれが何なのか（表題・説明）はメモリ上にしか無いので
        # ここで引き継ぐ。これが無いと証跡から「実行前」が丸ごと抜ける
        self.before_images = before_images if before_images is not None else []

    # --- ログオフセットの変換 ------------------------------------------------
    @staticmethod
    def encode_offsets(offsets):
        # type: (Dict[Path, Any]) -> List[Dict[str, Any]]
        """logs.snapshot_offsets() の戻り値を JSON にできる形へ。"""
        encoded = []
        for path, value in offsets.items():
            size, head_len, head_sig = value
            encoded.append({
                "path": str(path),
                "size": int(size),
                "head_len": int(head_len),
                "head_signature": str(head_sig),
            })
        return sorted(encoded, key=lambda item: item["path"])

    def decode_offsets(self):
        # type: () -> Dict[Path, Any]
        """logs.collect() が受け取れる形へ戻す。"""
        restored = {}
        for item in self.log_offsets:
            restored[Path(item["path"])] = (
                int(item["size"]), int(item["head_len"]), str(item["head_signature"]))
        return restored

    # --- 入出力 --------------------------------------------------------------
    def to_dict(self):
        # type: () -> Dict[str, Any]
        return {
            "format_version": FORMAT_VERSION,
            "session_id": self.session_id,
            "case_id": self.case_id,
            "case_source": self.case_source,
            "config_path": self.config_path,
            "offline": bool(self.offline),
            "base_date": self.base_date,
            "batch_name": self.batch_name,
            "log_dir": self.log_dir,
            "phase": self.phase,
            "before_started_at": _dumps_time(self.before_started_at),
            "before_finished_at": _dumps_time(self.before_finished_at),
            "log_offsets": self.log_offsets,
            "snapshot_tables": self.snapshot_tables,
            "before_images": self.before_images,
        }

    @staticmethod
    def encode_images(images, session_dir):
        # type: (List[Any], Path) -> List[Dict[str, Any]]
        """ImageEvidence を JSON にできる形へ。パスは session フォルダからの相対。

        相対にしておくと、証跡フォルダごと別の場所へコピーしても読める。
        """
        encoded = []
        for image in images:
            try:
                rel = str(Path(image.path).relative_to(session_dir)).replace("\\", "/")
            except ValueError:
                rel = str(image.path)
            encoded.append({"title": image.title, "path": rel, "caption": image.caption})
        return encoded

    def decode_images(self, session_dir):
        # type: (Path) -> List[Any]
        """ImageEvidence へ戻す。実体が無いものは黙って落とさず除外して数を残す。"""
        from .models import ImageEvidence

        restored = []
        for item in self.before_images:
            path = Path(item["path"])
            if not path.is_absolute():
                path = session_dir / path
            restored.append(ImageEvidence(
                title=str(item.get("title") or ""),
                path=path,
                caption=str(item.get("caption") or "")))
        return restored

    def save(self, session_dir):
        # type: (Path) -> Path
        session_dir.mkdir(parents=True, exist_ok=True)
        path = session_dir / "session.json"
        with path.open("w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2, sort_keys=True)
        return path

    @classmethod
    def load(cls, session_dir):
        # type: (Path) -> ManualSession
        """session.json を読む。壊れていたら推測せず ConfigError。"""
        path = session_dir / "session.json"
        if not path.exists():
            raise ConfigError(
                "手動実施の作業情報が見つかりません: %s\n"
                "  先に  python -m autotest manual --case <ID> --phase before  を実行してください。"
                % path)
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except ValueError as exc:
            raise ConfigError("%s を読めませんでした（壊れている可能性）: %s" % (path, exc))

        version = data.get("format_version")
        if version != FORMAT_VERSION:
            raise ConfigError(
                "%s の書式バージョンが違います（ファイル: %r / 想定: %d）。\n"
                "  ツールを更新した場合は、この手動ケースを before からやり直してください。"
                % (path, version, FORMAT_VERSION))

        for key in ("session_id", "case_id", "base_date", "phase"):
            if not data.get(key):
                raise ConfigError("%s に必須項目 %s がありません。" % (path, key))

        return cls(
            session_id=str(data["session_id"]),
            case_id=str(data["case_id"]),
            case_source=str(data.get("case_source") or ""),
            config_path=str(data.get("config_path") or ""),
            offline=bool(data.get("offline", False)),
            base_date=str(data["base_date"]),
            batch_name=data.get("batch_name"),
            log_dir=str(data.get("log_dir") or ""),
            phase=str(data["phase"]),
            before_started_at=_loads_time(data.get("before_started_at")),
            before_finished_at=_loads_time(data.get("before_finished_at")),
            log_offsets=list(data.get("log_offsets") or []),
            snapshot_tables=list(data.get("snapshot_tables") or []),
            before_images=list(data.get("before_images") or []),
        )


def session_prefix(case_id):
    # type: (str) -> str
    return "manual_%s_" % case_id


def new_session_id(case_id, started):
    # type: (str, datetime) -> str
    """run_id と同じ粒度（ミリ秒まで）。同一秒に 2 回始めても衝突しない。"""
    stamp = started.strftime("%Y%m%d_%H%M%S") + "_%03d" % (started.microsecond // 1000)
    return session_prefix(case_id) + stamp


def find_sessions(out_dir, case_id, only_open=True):
    # type: (Path, str, bool) -> List[Path]
    """このケースの手動作業フォルダを新しい順に返す。

    only_open=True なら「before は済んだが after がまだ」のものだけ。
    """
    if not out_dir.is_dir():
        return []
    found = []
    for path in sorted(out_dir.glob(session_prefix(case_id) + "*"), reverse=True):
        if not (path / "session.json").exists():
            continue
        if only_open:
            try:
                if ManualSession.load(path).phase != PHASE_BEFORE_DONE:
                    continue
            except ConfigError:
                # 壊れた session は「開いている」とはみなさない。
                # --session で明示指定されれば load 側がエラーを出す
                continue
        found.append(path)
    return found
