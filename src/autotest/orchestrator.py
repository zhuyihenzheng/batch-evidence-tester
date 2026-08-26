"""1 テストケースのライフサイクル制御。

    前処理 → 実行前スナップショット → .exe 実行 → 成果回収 → 実行後スナップショット
    → 証跡画像生成 → 判定

順序の要点:
  - 成果ファイルの回収は判定より先。batch の再実行や後続ケースで消えるため。
  - ログのオフセット取得は .exe 起動の直前。増分だけを切り出すため。
  - フォルダ一覧は「投入直後」と「実行後」の 2 時点を撮る。差が証跡になる。
"""

import re
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Tuple

from . import compare, db as db_mod, fsops, logs, render, screenshot
from .config import ConfigError, Settings, TestCase
from .models import (NG, OK, REVIEW, SKIP, CaseResult, CheckResult, ExecutionInfo,
                     ImageEvidence, Table)
from .runner import _resolve_exe, run_batch

# {batch_start} / {batch_start:%Y%m%d} の両方を受ける
_BATCH_START_PLACEHOLDER = re.compile(r"\{batch_start(?::(?P<fmt>[^}]+))?\}")


def preflight_case(settings: Settings, case: TestCase) -> List[str]:
    """破壊的操作（フォルダクリア・SQL 実行）の前に行う静的検査。

    「フォルダを空にして DB をリセットした後で exe が無いことに気づく」
    という事故を防ぐため、実行に必要な資材がすべて揃っているかを先に確認する。
    validate コマンドも同じ関数を使う。戻り値は問題の一覧（空なら合格）。
    """
    problems: List[str] = []

    # --- batch 定義と exe -------------------------------------------------
    batch_name = case.execute.get("batch")
    try:
        batch = settings.batch_profile(batch_name)
    except ConfigError as exc:
        problems.append(str(exc))
        batch = None

    if batch is not None:
        exe_raw = batch.get("exe_path")
        if not exe_raw:
            problems.append("batch の exe_path が未設定です（batch 名: %s）" % (batch_name or "既定"))
        else:
            exe = _resolve_exe(str(exe_raw), settings.project_root)
            if not exe.exists():
                problems.append("実行ファイルが見つかりません: %s" % exe)
        working_dir = batch.get("working_dir")
        if working_dir:
            wd = Path(str(working_dir))
            if not wd.is_absolute():
                wd = (settings.project_root / wd).resolve()
            if not wd.is_dir():
                problems.append("working_dir が存在しません: %s" % wd)

    aliases = settings.path_aliases

    # --- manual ケースで使えない機能 -----------------------------------------
    # replace_files と db_lock は「run() の finally で必ず元に戻す」前提の機能。
    # manual は before と after が別プロセスなので、その finally が存在しない。
    # 中途半端に対応すると「設定を差し替えたまま」「ロックを握ったまま」で
    # 終わる経路ができるため、定義の段階で断る。
    if case.is_manual:
        for key, what in (("replace_files", "設定ファイルの差し替え"),
                          ("db_lock", "DB ロックの保持")):
            if (case.setup or {}).get(key):
                problems.append(
                    "setup.%s は mode: manual のケースでは使えません（%s は自動実行の "
                    "後始末に依存しており、手動の 2 段階実行では元に戻す担保が無いため）。"
                    "手で配置・解除してください。" % (key, what))

    # --- clean_dirs / remove_dirs: 論理名の存在と、解決後パスの安全性 --------
    for key in ("clean_dirs", "remove_dirs"):
        for spec in (case.setup or {}).get(key, []):
            alias = spec["alias"] if isinstance(spec, dict) else spec
            if alias not in aliases:
                problems.append("setup.%s の論理名が paths に未定義です: %s" % (key, alias))
                continue
            try:
                fsops.assert_safe_to_clear(aliases[alias], settings.project_root)
            except ConfigError as exc:
                problems.append(str(exc))

            # 配下に別の論理名がある場合、そのフォルダごと消える構成になっていないか。
            # 例: error_dir=.../Receive の下に backup_dir=.../Receive/Backup がある
            target = aliases[alias]
            nested = [name for name, path in aliases.items()
                      if name != alias and path != target
                      and fsops._is_relative_to(path, target)]
            if nested and key == "remove_dirs":
                problems.append(
                    "setup.remove_dirs の %s (%s) の配下に別の論理名があります: %s。"
                    "フォルダごと削除するとこれらも消えます。" % (alias, target, nested)
                )

    # remove_dirs で消すフォルダへファイル・フォルダを投入すると、投入時に再作成されて
    # 「フォルダが無い」状態を作れない。設定の矛盾として先に弾く
    removing = set((case.setup or {}).get("remove_dirs", []))
    for spec in (case.setup or {}).get("input_files", []):
        dest = spec.get("dest_dir", "input_dir")
        if dest in removing:
            problems.append(
                "setup.remove_dirs に %s があるのに、同じフォルダへ input_files を投入しています。"
                "投入時にフォルダが再作成されるため「フォルダ無し」の再現になりません。" % dest
            )

    # --- setup の資材 -------------------------------------------------------
    for sql_spec in (case.setup or {}).get("sql", []):
        if isinstance(sql_spec, dict) and "file" in sql_spec:
            sql_path = settings.project_root / sql_spec["file"]
            if not sql_path.exists():
                problems.append("セットアップ SQL がありません: %s" % sql_path)

    for spec in (case.setup or {}).get("input_files", []):
        src_raw = spec.get("src")
        if not src_raw:
            problems.append("setup.input_files に src がありません: %s" % spec)
            continue
        src = Path(src_raw)
        if not src.is_absolute():
            src = case.dir / src
        if not src.exists():
            problems.append("投入元がありません: %s" % src)
        elif src.is_dir() and spec.get("corrupt"):
            problems.append("フォルダの投入には corrupt を指定できません: %s" % src)
        dest_alias = spec.get("dest_dir", "input_dir")
        if dest_alias not in aliases:
            problems.append("input_files.dest_dir の論理名が paths に未定義です: %s" % dest_alias)

    # --- replace_files（batch 設定ファイルの差し替え）-------------------------
    for spec in (case.setup or {}).get("replace_files", []):
        src_raw = spec.get("src")
        if not src_raw:
            problems.append("setup.replace_files に src がありません: %s" % spec)
            continue
        src = Path(src_raw)
        if not src.is_absolute():
            src = case.dir / src
        if not src.exists():
            problems.append("差し替え元ファイルがありません: %s" % src)
        dest_alias = spec.get("dest_dir")
        if not dest_alias:
            problems.append("setup.replace_files に dest_dir がありません: %s" % spec)
        elif dest_alias not in aliases:
            problems.append("replace_files.dest_dir の論理名が paths に未定義です: %s" % dest_alias)
        # 差し替え先を clean_dirs で消すと、退避前に元ファイルが失われる
        cleaning = {(c["alias"] if isinstance(c, dict) else c)
                    for c in (case.setup or {}).get("clean_dirs", [])}
        if dest_alias in cleaning:
            problems.append(
                "setup.clean_dirs に %s があるため、差し替え先のフォルダが空にされます。"
                "batch 本体や元の設定ファイルまで消えるので、この論理名は clean_dirs から外してください。"
                % dest_alias)

    # --- collect / folder_evidence ------------------------------------------
    for spec in (case.collect or {}).get("files", []):
        if spec.get("dir", "output_dir") not in aliases:
            problems.append("collect.files.dir の論理名が paths に未定義です: %s" % spec.get("dir"))
    for alias in (case.collect or {}).get("folder_evidence", []):
        if alias not in aliases:
            problems.append("collect.folder_evidence の論理名が paths に未定義です: %s" % alias)

    # --- 期待値ファイル -------------------------------------------------------
    # manual: true は「人が証跡を見て判定する」ため expected が無くてよい。
    # 無条件に必須にすると、その運用が preflight で弾かれて成立しない。
    for item in (case.assertions or {}).get("db", []):
        if not item.get("expected"):
            if not item.get("manual"):
                problems.append(
                    "assert.db に expected がありません: %s"
                    "（人が目視で判定する場合は manual: true を指定してください）" % item.get("table"))
            continue
        expected = settings.project_root / item["expected"]
        if not expected.exists():
            problems.append("期待値 CSV がありません: %s" % expected)

    # DB 判定は実行後スナップショットを入力にする。ここを紐付け忘れると
    # 実行後に初めて NG になるため、validate の時点で具体的に案内する。
    snapshot_names = {
        str(item.get("name") or item.get("table") or "")
        for item in (case.snapshot or {}).get("tables", [])
        if isinstance(item, dict)
    }
    for item in (case.assertions or {}).get("db", []):
        table_name = str(item.get("table") or "")
        if table_name and table_name not in snapshot_names:
            problems.append(
                "assert.db の %s が snapshot.tables にありません。"
                "実行前後の証跡を採るため snapshot.tables に同じ名前で定義してください。"
                % table_name)

    for item in (case.assertions or {}).get("files", []):
        if "expected" in item:
            expected = settings.project_root / item["expected"]
            if not expected.exists():
                problems.append("期待値ファイルがありません: %s" % expected)
        actual_spec = item.get("actual") or item.get("exists") or {}
        if actual_spec and actual_spec.get("dir", "output_dir") not in aliases:
            problems.append("assert.files の dir 論理名が paths に未定義です: %s" % actual_spec.get("dir"))

        # 期待値を持たない目視確認は、実ファイルの中身が証跡に無ければ判定不能。
        # 同じ採取指定（dir / pattern）に preview: true があることを必須にする。
        if item.get("manual") is True and "actual" in item:
            actual = item.get("actual") or {}
            actual_dir = actual.get("dir", "output_dir")
            actual_pattern = actual.get("pattern", "*")
            matching_collect = [
                c for c in (case.collect or {}).get("files", [])
                if isinstance(c, dict)
                and c.get("dir", "output_dir") == actual_dir
                and c.get("pattern", "*") == actual_pattern
                and c.get("preview") is True
            ]
            if not matching_collect:
                problems.append(
                    "assert.files の目視確認 (%s / %s) に対応する証跡がありません。"
                    "collect.files に同じ dir / pattern と preview: true を指定してください。"
                    % (actual_dir, actual_pattern))

    return problems


def _force_review(check: CheckResult) -> CheckResult:
    """手動実施ケースの判定項目を「要確認」に落とす。

    人が手で動かした実行の自動比較は、あくまで参考値。ここで OK を通すと
    「人が確認していないのに合格」が成立してしまう。NG はそのまま NG
    （確認を待たずに問題として扱う）。SKIP もそのまま。
    """
    if check.verdict != OK:
        return check
    check.verdict = REVIEW
    check.detail = "自動比較は一致（%s）。手動実行のため、最終判定は目視で行ってください。" % check.detail
    return check


def _apply_manual(check: CheckResult, manual: bool) -> CheckResult:
    """manual 指定なら、自動判定が OK でも確定させず「要確認」にする。

    自動比較の結果は detail に残すので、人はそれを踏まえて判断できる。
    NG はそのまま NG（人の確認を待たずに問題として扱う）。
    """
    if not manual or check.verdict != OK:
        return check
    check.verdict = REVIEW
    check.detail = "自動比較は一致（%s）。最終判定は目視で行ってください。" % check.detail
    return check


class CaseRunner:
    def __init__(self, settings: Settings, run_dir: Path, offline: bool = False,
                 dry_run: bool = False, progress=None, default_date=None,
                 keep_env: bool = False):
        self.settings = settings
        self._default_date = default_date
        self.run_dir = run_dir
        self.offline = offline
        self.dry_run = dry_run
        # 調査用。失敗した状態を保全するため、後始末を行わない
        self.keep_env = keep_env
        # 各工程の開始を通知するコールバック。処理が止まったとき、
        # どの工程で止まっているかを外から見えるようにするためのもの
        self._progress = progress or (lambda step: None)
        # batch 起動直前に DB から取る基準時刻（{batch_start} の展開に使う）
        self._batch_start_mark = None
        # 差し替え中のファイル（ケース終了時に元へ戻す）
        self._replaced_files = []

    def _step(self, name: str) -> None:
        self._progress(name)

    # ------------------------------------------------------------------
    def run(self, case: TestCase) -> CaseResult:
        result = CaseResult(case_id=case.case_id, name=case.name, description=case.description, tags=case.tags)
        evidence_dir = self.run_dir / "evidence" / case.case_id
        artifact_dir = self.run_dir / "artifacts" / case.case_id
        renderer = render.Renderer(self.settings.evidence, evidence_dir)

        # このケースの基準日を確定する。ケースが execute.date を持つならそれを、
        # 無ければ実行単位の既定（--date か本日）を使う。パス・引数・パターンの
        # {date} は以降この日付で展開される。ケースは直列実行のため競合しない。
        self.settings.set_base_date(case.execute.get("date") or self._default_date)

        # 破壊的操作（クリア・SQL）より前に、実行に必要な資材を全部確認する。
        # ここで弾けば「フォルダを空にした後で exe が無いと判明」が起きない。
        problems = preflight_case(self.settings, case)
        if problems:
            result.fatal_error = "事前チェック（preflight）で不合格。破壊的操作は行っていません:\n  - " + "\n  - ".join(problems)
            return result

        client: Optional[db_mod.DbClient] = None
        setup_started = False
        try:
            self._step("DB接続")
            client = db_mod.create_client(self.settings, self.offline, case.case_id, dry_run=self.dry_run)

            self._step("前処理（フォルダ準備・SQL・データ投入）")
            # teardown が必要かどうかの判断に使う。setup に入った時点で
            # 環境や DB が変更されている可能性があるため、以降は必ず後始末する
            setup_started = True
            self._setup(case, client)

            self._step("フォルダ撮影（実行前）")
            folder_before = self._capture_folders(renderer, "実行前", case)

            # {batch_start} を使うケースだけ DB から基準時刻を取る。
            # 使わないケースで毎回問い合わせるのは無駄だし、失敗の種を増やすだけ。
            self._batch_start_mark = self._resolve_batch_start(case, client)

            self._step("DBスナップショット（実行前）")
            self._snapshot_db(case, client, result, phase="before")

            batch_name = case.execute.get("batch")
            log_dir = self._log_dir_for(batch_name)

            self._step("ログ位置の記録")
            offsets = logs.snapshot_offsets(self.settings, log_dir)

            # DB 更新失敗の再現。batch 実行中だけ別接続でロックを保持する。
            # with を抜けるときに必ずロールバックするので、失敗しても
            # ロックが残って後続が止まることはない。
            lock_sql = (case.setup or {}).get("db_lock")
            lock_ctx = None
            if lock_sql and not self.dry_run and not self.offline:
                self._step("DBロック取得（更新失敗の再現）")
                lock_ctx = db_mod.LockHolder(self.settings, self._expand_sql(str(lock_sql)))
                lock_ctx.__enter__()

            self._step("batch実行")
            result.execution = run_batch(
                self.settings,
                case.execute.get("args", []),
                dry_run=self.dry_run,
                batch_name=batch_name,
                on_progress=self._progress,   # 実行中の経過を run.log と画面へ
            )

            # .exe が終了してもログの書き込みが終わっているとは限らない。
            # 落ち着くまで待ってから読む（既に落ち着いていれば即座に進む）
            batch_cfg = self.settings.batch_profile(batch_name)
            waited = logs.wait_until_stable(
                self.settings, log_dir,
                max_wait_sec=float(batch_cfg.get("log_settle_sec", 5)),
                min_wait_sec=float(batch_cfg.get("log_settle_min_sec", 1)),
                stable_for_sec=float(batch_cfg.get("log_settle_stable_sec", 1)),
            )
            if waited >= 0.5:
                self._step("ログ書き込み待ち %.1f 秒" % waited)

            if lock_ctx is not None:
                # batch が終わったら即座に解放する。判定や証跡採取の間
                # 保持し続ける必要はなく、長く持つほど巻き添えが増える
                self._step("DBロック解放")
                lock_ctx.__exit__(None, None, None)
                lock_ctx = None

            self._step("ログ収集")
            log_slices = logs.collect(
                self.settings, offsets,
                result.execution.started_at, result.execution.finished_at,
                log_dir=log_dir,
            )

            self._step("成果ファイル回収")
            result.saved_artifacts = fsops.collect_artifacts(
                self.settings, case.collect.get("files", []), artifact_dir
            )

            self._step("DBスナップショット（実行後）")
            self._snapshot_db(case, client, result, phase="after")

            self._step("フォルダ撮影（実行後）")
            folder_after = self._capture_folders(renderer, "実行後", case)

            self._step("証跡生成")
            result.log_tables = self._build_log_tables(log_slices, case, log_dir)
            result.images.extend(folder_before)
            result.images.extend(folder_after)
            result.images.extend(self._render_file_previews(renderer, case, result))
            result.images.extend(self._render_replaced_configs(renderer, case))

            self._step("判定")
            if self.dry_run:
                # dry-run は batch も DB も動かしていないので、判定しても意味がない。
                # ここで NG を出すと「設定が悪いのか、判定が落ちたのか」が区別できず
                # 誤解を招くため、SKIP として記録するに留める
                result.checks = [
                    CheckResult("判定（dry-run のためスキップ）", "dry_run", SKIP,
                                "batch も DB も実行していないため判定していません。"
                                "前処理・フォルダ操作・設定の確認のみ行いました。")
                ]
            else:
                result.checks = self._assert_all(case, client, result, log_slices)

        except Exception as exc:  # 1 ケースの失敗で全体を止めない
            result.fatal_error = f"{type(exc).__name__}: {exc}\n{traceback.format_exc(limit=3)}"
        finally:
            # 例外で抜けた場合もロックを必ず解放する。保持したままだと
            # 後続ケースも他の作業も全部ロック待ちで止まる
            if locals().get("lock_ctx") is not None:
                try:
                    lock_ctx.__exit__(None, None, None)
                except Exception:
                    pass

            # 後処理（teardown）は finally で行う。証跡採取や判定の途中で
            # 例外が出た場合でも、setup が入れたテストデータ・制約違反データ・
            # 権限変更を DB に残したままにしないため。
            if setup_started and client is not None and not self.keep_env:
                try:
                    teardown_check = self._teardown(case, client)
                    if teardown_check is not None:
                        result.checks.append(teardown_check)
                except Exception as exc:
                    result.checks.append(CheckResult(
                        "後処理（teardown）", "teardown", NG,
                        "teardown 自体が失敗しました（環境が変更されたまま残っています）: %s" % exc))

            # 差し替えたファイルは、途中で落ちた場合でも必ず元に戻す。
            # 戻し損ねると次のケースや手動実行が壊れた設定で動いてしまう
            if self.keep_env:
                # 調査用に差し替えたままにする。何を戻していないかは残す
                if self._replaced_files:
                    result.checks.append(CheckResult(
                        "設定ファイルの復元", "teardown", SKIP,
                        "--keep-env のため復元していません。調査後に手動で戻してください: "
                        + ", ".join(str(r.dest) for r in self._replaced_files)))
                restore_problems = []
            else:
                restore_problems = fsops.restore_files(getattr(self, "_replaced_files", []))
            self._replaced_files = []
            if restore_problems:
                result.checks.append(CheckResult(
                    "設定ファイルの復元", "teardown", NG,
                    "差し替えたファイルを元に戻せませんでした（環境が変更されたままです）: "
                    + " / ".join(restore_problems)))
            if client:
                client.close()

        return result

    # ==================================================================
    # 手動実施（mode: manual）
    #   run() を before / after の 2 つに割ったもの。工程そのものは
    #   run() と同じヘルパを呼ぶ —— ライフサイクルの知識を 2 か所に
    #   持たせると、片方だけ直して静かにずれる。
    # ==================================================================
    def run_manual_before(self, case: TestCase, session_dir: Path):
        """前処理と実行前の証跡採取まで行い、引き継ぎ情報を返す。

        この後、人が手で batch を動かす。戻り値の ManualSession を
        session_dir に保存しておけば、別プロセスの after が続きを行える。
        """
        from . import manual as manual_mod

        self.settings.set_base_date(case.execute.get("date") or self._default_date)

        problems = preflight_case(self.settings, case)
        if problems:
            raise ConfigError(
                "事前チェック（preflight）で不合格。破壊的操作は行っていません:\n  - "
                + "\n  - ".join(problems))

        started = datetime.now()
        renderer = render.Renderer(self.settings.evidence, session_dir / "evidence" / case.case_id)
        client = db_mod.create_client(self.settings, self.offline, case.case_id, dry_run=False)
        try:
            self._step("前処理（フォルダ準備・SQL・データ投入）")
            self._setup(case, client)

            self._step("フォルダ撮影（実行前）")
            images = self._capture_folders(renderer, "実行前", case)

            self._batch_start_mark = self._resolve_batch_start(case, client)

            self._step("DBスナップショット（実行前）")
            result = CaseResult(case_id=case.case_id, name=case.name)
            self._snapshot_db(case, client, result, phase="before")
            # 実行前スナップショットは CSV として残す。after は別プロセスなので
            # メモリ上の Table を引き継げない。既存の期待値 CSV と同じ書式なので
            # db_mod.read_expected_csv でそのまま読み戻せる
            self._save_snapshot_csvs(result.db_before, session_dir / "before")

            batch_name = case.execute.get("batch")
            log_dir = self._log_dir_for(batch_name)
            self._step("ログ位置の記録")
            offsets = logs.snapshot_offsets(self.settings, log_dir)
        finally:
            client.close()

        # 撮影済みの画像は after プロセスが証跡へ組み込む。画像ファイル自体は
        # session フォルダに残るが、表題と説明はメモリ上にしか無いので
        # session.json へ引き継ぐ。ここを落とすと証跡から「実行前」が丸ごと消える
        return manual_mod.ManualSession(
            session_id=session_dir.name,
            case_id=case.case_id,
            case_source=str(case.source),
            base_date=self.settings.base_date.strftime("%Y%m%d"),
            batch_name=batch_name,
            log_dir=str(log_dir),
            offline=self.offline,
            phase=manual_mod.PHASE_BEFORE_DONE,
            before_started_at=started,
            before_finished_at=datetime.now(),
            log_offsets=manual_mod.ManualSession.encode_offsets(offsets),
            snapshot_tables=sorted(result.db_before),
            before_images=manual_mod.ManualSession.encode_images(images, session_dir),
        )

    def run_manual_after(self, case: TestCase, session, session_dir: Path) -> CaseResult:
        """人が batch を動かした後の証跡採取と判定。

        判定結果は必ず「要確認」で頭打ちにする。人が手で動かした以上、
        自動比較の一致は参考値であって合格の根拠にはならない。
        """
        from . import manual as manual_mod

        # 基準日は session から復元する。当日の日付で展開すると、
        # before と after が日をまたいだときに {date} を含むパターンが
        # 静かに外れて「出力が無い」と誤判定する
        self.settings.set_base_date(session.base_date)

        result = CaseResult(case_id=case.case_id, name=case.name,
                            description=case.description, tags=case.tags)
        evidence_dir = session_dir / "evidence" / case.case_id
        # before プロセスが同じフォルダへ描いた画像の続きから番号を振る
        already = len(list(evidence_dir.glob("*.png"))) if evidence_dir.is_dir() else 0
        renderer = render.Renderer(self.settings.evidence, evidence_dir, start_seq=already)
        collected_at = datetime.now()

        client = db_mod.create_client(self.settings, self.offline, case.case_id, dry_run=False)
        try:
            # 実行前スナップショットを CSV から復元する（before プロセスの成果）。
            # 取得元の SQL も付け直す —— 実行後だけに SQL が載っていて実行前に
            # 無いと、証跡としては「この表が何を問い合わせた結果か」が片方だけ
            # 不明な、読み手に不親切な状態になる
            sql_by_table = {}
            for spec in case.snapshot.get("tables", []):
                table_name = str(spec.get("name") or spec.get("table") or "")
                if table_name:
                    sql_by_table[table_name] = str(
                        spec.get("sql") or "SELECT * FROM %s" % table_name).strip()
            for name in session.snapshot_tables:
                path = session_dir / "before" / ("db_%s.csv" % name)
                if path.exists():
                    table = db_mod.read_expected_csv(path)
                    table.title = name
                    if name in sql_by_table:
                        table.note = "SQL: %s" % self._expand_sql(sql_by_table[name])
                    result.db_before[name] = table

            log_dir = Path(session.log_dir) if session.log_dir else self._log_dir_for(session.batch_name)

            self._step("ログ書き込み待ち")
            batch_cfg = self.settings.batch_profile(session.batch_name)
            logs.wait_until_stable(
                self.settings, log_dir,
                max_wait_sec=float(batch_cfg.get("log_settle_sec", 5)),
                min_wait_sec=float(batch_cfg.get("log_settle_min_sec", 1)),
                stable_for_sec=float(batch_cfg.get("log_settle_stable_sec", 1)),
            )

            self._step("ログ収集")
            log_slices = logs.collect(
                self.settings, session.decode_offsets(),
                session.before_finished_at, collected_at, log_dir=log_dir)

            self._step("成果ファイル回収")
            result.saved_artifacts = fsops.collect_artifacts(
                self.settings, case.collect.get("files", []),
                session_dir / "artifacts" / case.case_id)

            self._step("DBスナップショット（実行後）")
            self._snapshot_db(case, client, result, phase="after")

            self._step("フォルダ撮影（実行後）")
            folder_after = self._capture_folders(renderer, "実行後", case)

            self._step("証跡生成")
            result.log_tables = self._build_log_tables(log_slices, case, log_dir)
            # 実行前の画像を先に置く。証跡は「実行前 → 実行後」の並びで読ませる
            result.images.extend(session.decode_images(session_dir))
            result.images.extend(folder_after)
            result.images.extend(self._render_file_previews(renderer, case, result))

            result.execution = ExecutionInfo(
                command="（手動実行）batch は利用者が手で起動",
                started_at=session.before_finished_at,
                finished_at=collected_at,
            )

            self._step("判定")
            converted = []
            for check in self._assert_all(case, client, result, log_slices):
                if check.category == "exit_code":
                    # 手動起動では終了コードを取得できない。利用者の申告を
                    # 受け付けると、検証できない値で合格が出せてしまう
                    converted.append(CheckResult(
                        "終了コード確認", "exit_code", SKIP,
                        "手動実行のため終了コードは取得していません。"
                        "異常終了したかどうかはログと DB の状態で判定してください。"))
                else:
                    converted.append(_force_review(check))
            result.checks = [self._manual_banner()] + converted
        finally:
            try:
                teardown_check = self._teardown(case, client)
                if teardown_check is not None:
                    result.checks.append(teardown_check)
            except Exception as exc:
                result.checks.append(CheckResult(
                    "後処理（teardown）", "teardown", NG,
                    "teardown 自体が失敗しました（環境が変更されたまま残っています）: %s" % exc))
            client.close()

        session.phase = manual_mod.PHASE_FINISHED
        return result

    @staticmethod
    def _manual_banner() -> CheckResult:
        """手動ケースであることを判定明細の先頭に必ず置く。

        これがあるとケースの verdict は決して OK にならない（REVIEW が 1 件でも
        あれば要確認）。判定項目が 1 つも無い手動ケースでも「確認待ち」で残る。
        """
        return CheckResult(
            "実行方式", "manual", REVIEW,
            "手動実行ケース。自動比較の結果は参考値です。証跡を確認して最終判定を記入してください。")

    def _save_snapshot_csvs(self, tables, out_dir: Path) -> None:
        """DB スナップショットを CSV として保存する。

        値は db.to_text() 済みの文字列なので、そのまま書けば情報は落ちない
        （整形や丸めは一切しない、というツール全体の原則と同じ）。
        """
        import csv

        out_dir.mkdir(parents=True, exist_ok=True)
        for name, table in tables.items():
            path = out_dir / ("db_%s.csv" % name)
            with path.open("w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(table.columns)
                for row in table.rows:
                    writer.writerow(row)

    # ------------------------------------------------------------------
    def _setup(self, case: TestCase, client: db_mod.DbClient) -> None:
        setup = case.setup or {}

        # フォルダの自動作成は「このケースが実際に使う論理名」だけに限定する。
        # paths 全部を作ると、無関係な batch のフォルダまで生やしてしまい、
        # その中にアクセスできないネットワーク共有があると巻き添えで失敗する。
        # dry-run では 1 つも作らない（何も触らないのが dry-run の約束）。
        if not self.dry_run:
            removing = set(setup.get("remove_dirs", []))
            fsops.ensure_dirs(
                self.settings,
                [a for a in self._aliases_used_by(case) if a not in removing])

        for spec in setup.get("clean_dirs", []):
            # 文字列（論理名だけ）と辞書（exclude 付き）の両方を受ける
            if isinstance(spec, dict):
                fsops.clear_dir(self.settings, spec["alias"], dry_run=self.dry_run,
                                exclude=spec.get("exclude"))
            else:
                fsops.clear_dir(self.settings, spec, dry_run=self.dry_run)

        # フォルダごと削除。batch が「出力先が無い」ときにどう振る舞うかを試す
        for alias in setup.get("remove_dirs", []):
            fsops.remove_dir(self.settings, alias, dry_run=self.dry_run)

        for sql_spec in setup.get("sql", []):
            if isinstance(sql_spec, str):
                sql_text = sql_spec
            elif "file" in sql_spec:
                sql_path = self.settings.project_root / sql_spec["file"]
                if not sql_path.exists():
                    raise ConfigError(f"セットアップ SQL が見つかりません: {sql_path}")
                sql_text = sql_path.read_text(encoding="utf-8-sig")
            else:
                sql_text = sql_spec.get("inline", "")
            if sql_text.strip() and not self.dry_run:
                client.execute_script(sql_text)

        fsops.put_input_files(self.settings, case.dir, setup.get("input_files", []), dry_run=self.dry_run)

        # batch の設定ファイル等をケース用に差し替える。元ファイルは退避し、
        # ケース終了時に必ず戻す（環境を変更したまま終わらせない）
        self._replaced_files = fsops.replace_files(
            self.settings, case.dir, setup.get("replace_files", []),
            backup_root=self.run_dir / "backup" / case.case_id,
            dry_run=self.dry_run,
        )

    def _aliases_used_by(self, case: TestCase) -> List[str]:
        """このケースが参照する論理名だけを集める。

        paths に定義された全フォルダを作ると、他機能の batch のフォルダまで
        生成してしまう。到達できないネットワーク共有が含まれていると、
        関係のないケースがそれに引きずられて失敗する。
        """
        setup = case.setup or {}
        collect = case.collect or {}
        used = set()

        for key in ("clean_dirs", "remove_dirs"):
            for spec in setup.get(key, []):
                used.add(spec["alias"] if isinstance(spec, dict) else spec)
        for spec in setup.get("input_files", []):
            used.add(spec.get("dest_dir", "input_dir"))
        for spec in setup.get("replace_files", []):
            if spec.get("dest_dir"):
                used.add(spec["dest_dir"])
        for spec in collect.get("files", []):
            used.add(spec.get("dir", "output_dir"))
        used.update(collect.get("folder_evidence") or self.settings.folder_evidence.get("targets", []))
        for item in (case.assertions or {}).get("files", []):
            for spec in (item.get("actual"), item.get("exists")):
                if spec:
                    used.add(spec.get("dir", "output_dir"))
        # ログ出力先は batch 定義側で決まる
        log_alias = self.settings.batch_profile(case.execute.get("batch")).get("log_dir")
        used.add(log_alias or "log_dir")

        aliases = self.settings.path_aliases
        return [a for a in used if a in aliases]

    # ------------------------------------------------------------------
    def _teardown(self, case: TestCase, client: db_mod.DbClient) -> Optional[CheckResult]:
        """後処理。setup と同じ書式（sql / clean_dirs）を実行する。

        定義が無ければ何もしない（None）。失敗はケースの NG として記録する
        —— 後続ケースの前提が壊れている可能性を隠さないため。
        """
        teardown = case.teardown or {}
        if not teardown:
            return None

        try:
            for sql_spec in teardown.get("sql", []):
                if isinstance(sql_spec, str):
                    sql_text = sql_spec
                elif "file" in sql_spec:
                    sql_path = self.settings.project_root / sql_spec["file"]
                    if not sql_path.exists():
                        raise ConfigError(f"teardown SQL が見つかりません: {sql_path}")
                    sql_text = sql_path.read_text(encoding="utf-8-sig")
                else:
                    sql_text = sql_spec.get("inline", "")
                if sql_text.strip() and not self.dry_run:
                    client.execute_script(sql_text)

            for spec in teardown.get("clean_dirs", []):
                if isinstance(spec, dict):
                    fsops.clear_dir(self.settings, spec["alias"], dry_run=self.dry_run,
                                    exclude=spec.get("exclude"))
                else:
                    fsops.clear_dir(self.settings, spec, dry_run=self.dry_run)
        except Exception as exc:
            return CheckResult(
                "後処理（teardown）", "teardown", NG,
                f"teardown に失敗しました（後続ケースの前提が壊れている可能性）: {exc}",
            )
        return CheckResult("後処理（teardown）", "teardown", OK, "sql/clean_dirs を実行済み")

    # ------------------------------------------------------------------
    def _capture_folders(self, renderer: render.Renderer, phase: str, case: TestCase) -> List[ImageEvidence]:
        cfg = self.settings.folder_evidence
        # ケース側で collect.folder_evidence: [...] を書くと撮影対象を差し替えられる。
        # batch ごとに入出力フォルダが違う場合に使う。
        targets = case.collect.get("folder_evidence") or cfg.get("targets", [])
        unknown = [t for t in targets if t not in self.settings.path_aliases]
        if unknown:
            raise ConfigError(
                "%s の collect.folder_evidence に paths 未定義の論理名があります: %s / 定義済み: %s"
                % (case.case_id, unknown, sorted(self.settings.path_aliases))
            )

        images: List[ImageEvidence] = []
        for alias in targets:
            directory = self.settings.resolve_dir(alias)
            entries = fsops.list_dir(
                directory,
                exclude_patterns=cfg.get("exclude_patterns", []),
                recursive=bool(cfg.get("recursive", False)),
            )
            path, note = self._folder_image(renderer, directory, entries, alias, phase, int(cfg.get("max_entries", 40)))
            n_dir = sum(1 for e in entries if e.is_dir)
            caption = f"{directory}  —  ファイル {len(entries) - n_dir} 件"
            if n_dir:
                caption += f" / フォルダ {n_dir} 件"
            if note:
                caption += f"  ※ {note}"
            images.append(
                ImageEvidence(
                    title=f"フォルダ確認: {alias}（{phase}）",
                    path=path,
                    caption=caption,
                )
            )
        return images

    def _folder_image(
        self,
        renderer: render.Renderer,
        directory: Path,
        entries: List[fsops.FileEntry],
        alias: str,
        phase: str,
        max_entries: int,
    ) -> "Tuple[Path, str]":
        """mode に応じて実画面キャプチャを試し、失敗したらコード描画へフォールバックする。

        戻り値は (画像パス, 注記)。フォールバックが起きた事実を証跡に残す
        —— screen モードの画像は前面ウィンドウを撮るため、黙って差し替わると
        「実画面」と「描画」の区別がつかず証跡の性質が変わってしまうため。
        """
        mode = self.settings.evidence.get("mode", "render")
        if mode in ("screen", "both"):
            try:
                out = renderer.out_dir / f"screen_{alias}_{phase}.png"
                return screenshot.capture_explorer(directory, out), "実画面キャプチャ"
            except screenshot.ScreenCaptureUnavailable as exc:
                note = f"実画面キャプチャ不可（{exc}）のため描画方式にフォールバック"
            except Exception as exc:
                note = f"実画面キャプチャ失敗（{type(exc).__name__}）のため描画方式にフォールバック"
        else:
            note = ""
        return renderer.folder_listing(directory, entries, alias, phase, max_entries), note

    # ------------------------------------------------------------------
    def _snapshot_db(self, case: TestCase, client: db_mod.DbClient, result: CaseResult, phase: str) -> None:
        specs = case.snapshot.get("tables", [])
        if not specs:
            return
        if isinstance(client, db_mod.OfflineClient):
            client.set_phase(phase)

        # 判定用データは打ち切らない（表示行数の絞り込みは excel.py 側で行う）
        target = result.db_before if phase == "before" else result.db_after
        for spec in specs:
            resolved = dict(spec)
            if spec.get("sql"):
                resolved["sql"] = self._expand_sql(str(spec["sql"]))
            table = client.snapshot(resolved)
            target[table.title] = table

    def _resolve_batch_start(self, case: TestCase, client: db_mod.DbClient):
        """{batch_start} の基準時刻を決める。

        DB サーバの時計から取る。テスト機と DB でずれていると、batch が
        最初に更新した行を取りこぼす（＝異常の見逃し）ため。
        使っていないケースでは問い合わせ自体を行わない。
        """
        specs = case.snapshot.get("tables", [])
        if not any("{batch_start" in str(spec.get("sql", "")) for spec in specs):
            return None
        # DB サーバの時計を使う。テスト機との時計ずれで対象行を取りこぼさないため
        mark = db_mod.server_now(client)
        if mark is not None:
            return mark
        # 取得できない場合のみテスト機の時計で代用する。進んでいた場合に
        # 取りこぼさないよう、マージン分さかのぼる
        margin = int(self.settings.database.get("batch_start_margin_sec", 5))
        return datetime.now() - timedelta(seconds=margin)

    def _expand_sql(self, sql: str) -> str:
        """スナップショット SQL のプレースホルダを展開する。

        {batch_start}            … batch 起動直前に DB サーバから取得した時刻。
                                   ISO 8601（T 区切り）で埋め込む。datetime 列と
                                   比較する場合はこれをそのまま使う。
        {batch_start:%Y%m%d}     … 書式指定。日付が char/varchar 列に
                                   'yyyyMMdd' 等で入っている場合に合わせる。
        {date} 系                 … 業務日付（paths や引数と同じ規則）
        """
        sql = self.settings.expand(sql)
        if "{batch_start" not in sql:
            return sql

        mark = getattr(self, "_batch_start_mark", None)
        if mark is None:
            # DB から時刻を取れない場合（offline 等）はテスト機の時刻で代用する。
            # 時計ずれの影響を受けるため、実 DB では server_now を優先している。
            mark = datetime.now()

        def replace(m):
            fmt = m.group("fmt")
            return mark.strftime(fmt) if fmt else db_mod.format_sql_datetime(mark)

        return _BATCH_START_PLACEHOLDER.sub(replace, sql)

    # ------------------------------------------------------------------
    def _log_dir_for(self, batch_name: Optional[str]) -> Path:
        """この batch のログ出力先を決める。

        batches.<名前>.log_dir に paths の論理名を書くと、batch ごとに
        別のログフォルダを見に行ける。未指定なら paths.log_dir。
        """
        alias = self.settings.batch_profile(batch_name).get("log_dir")
        return self.settings.resolve_dir(alias or "log_dir")

    def _build_log_tables(self, slices: List[logs.LogSlice], case: TestCase, log_dir: Path) -> List[Table]:
        """ログを Excel のネイティブ表に変換する（画像化しない）。"""
        keywords = list((case.assertions.get("log") or {}).get("must_contain") or [])
        classifier = logs.LineClassifier(self.settings.log)

        if not slices:
            return [
                Table(
                    title="（増分なし）",
                    columns=["行", "内容"],
                    rows=[[1, "この実行で増加したログ行はありません"]],
                    note=str(log_dir),
                    mono_columns=[1],
                )
            ]
        max_lines = int(self.settings.log.get("max_lines_in_excel", 500))
        return [logs.to_table(sl, classifier, keywords, max_lines=max_lines) for sl in slices]

    # ------------------------------------------------------------------
    def _render_replaced_configs(self, renderer: render.Renderer, case: TestCase) -> List[ImageEvidence]:
        """差し替えた設定ファイルの中身を証跡に残す。

        「このケースはどの設定で動かしたのか」は後から確認できないと
        再現も原因調査もできないため、実際に配置した内容を記録する。
        """
        images: List[ImageEvidence] = []
        cfg = self.settings.evidence
        for spec in (case.setup or {}).get("replace_files", []):
            src = Path(spec["src"])
            if not src.is_absolute():
                src = case.dir / src
            if not src.exists():
                continue
            dest_dir = self.settings.resolve_dir(spec.get("dest_dir", ""))
            name = self.settings.expand(spec.get("name") or src.name)
            text = fsops.read_text_file(
                src, self.settings.log.get("encoding", "utf-8"),
                self.settings.log.get("encoding_fallbacks", ["cp932"]))
            paths = renderer.text_page(
                title=f"差し替えた設定: {name}",
                subtitle=str(dest_dir / name),
                lines=text.splitlines(),
                stem=f"config_{Path(name).stem}",
                max_lines=int(cfg.get("file_preview_lines", 40)),
                max_cols=int(cfg.get("file_preview_cols", 160)),
            )
            for i, path in enumerate(paths, start=1):
                images.append(ImageEvidence(
                    title=f"差し替えた設定: {name}" + (f" ({i}/{len(paths)})" if len(paths) > 1 else ""),
                    path=path,
                    caption=f"差し替え元: {src}  /  実行後に元の内容へ復元済み",
                ))
        return images

    def _render_file_previews(self, renderer: render.Renderer, case: TestCase, result: CaseResult) -> List[ImageEvidence]:
        """回収したファイルの中身を画像化する（preview: true の指定分のみ）。"""
        images: List[ImageEvidence] = []
        cfg = self.settings.evidence
        max_lines = int(cfg.get("file_preview_lines", 40))
        max_cols = int(cfg.get("file_preview_cols", 160))

        preview_aliases = {
            spec.get("dir", "output_dir") for spec in case.collect.get("files", []) if spec.get("preview")
        }
        if not preview_aliases:
            return images

        for saved in result.saved_artifacts:
            if saved.parent.name not in preview_aliases:
                continue
            text = fsops.read_text_file(
                saved,
                self.settings.log.get("encoding", "utf-8"),
                self.settings.log.get("encoding_fallbacks", ["cp932"]),
            )
            paths = renderer.text_page(
                title=f"ファイル内容: {saved.name}",
                subtitle=str(self.settings.resolve_dir(saved.parent.name) / saved.name),
                lines=text.splitlines(),
                stem=f"file_{saved.stem}",
                max_lines=max_lines,
                max_cols=max_cols,
            )
            for i, p in enumerate(paths, start=1):
                images.append(
                    ImageEvidence(
                        title=f"ファイル内容: {saved.name}" + (f" ({i}/{len(paths)})" if len(paths) > 1 else ""),
                        path=p,
                        caption=f"SHA-256: {fsops.file_sha256(saved)[:16]}...  /  {saved.stat().st_size:,} bytes",
                    )
                )
        return images

    # ------------------------------------------------------------------
    def _assert_all(
        self,
        case: TestCase,
        client: db_mod.DbClient,
        result: CaseResult,
        log_slices: List[logs.LogSlice],
    ) -> List[CheckResult]:
        spec = case.assertions or {}
        checks: List[CheckResult] = self._review_points(case)

        # assert は任意。証跡採取 + review だけのケースに「期待値なし」の
        # SKIP 行を混ぜず、YAML に書いた人工確認観点だけを Excel に並べる。
        if "exit_code" in spec:
            checks.append(self._assert_exit_code(spec, result))
        if "db" in spec:
            checks.extend(self._assert_db(spec, result))
        if "files" in spec:
            checks.extend(self._assert_files(spec))
        if "log" in spec:
            checks.append(self._assert_log(spec, log_slices))

        return [c for c in checks if c is not None]

    def _review_points(self, case: TestCase) -> List[CheckResult]:
        """YAML の review を、Excel で一項目ずつ確定する確認行へ変換する。"""
        checks: List[CheckResult] = []
        for content in case.review:
            text = str(content).strip()
            checks.append(CheckResult(
                text, "review", REVIEW, ""))
        return checks

    def _assert_exit_code(self, spec: dict, result: CaseResult) -> CheckResult:
        if "exit_code" not in spec:
            return CheckResult("終了コード確認", "exit_code", SKIP, "期待値の指定なし")
        exit_spec = spec["exit_code"]
        if isinstance(exit_spec, dict):
            expected = exit_spec["expected"]
            manual = exit_spec.get("manual") is True
        else:
            expected = exit_spec
            manual = False
        actual = result.execution.exit_code
        if result.execution.timed_out:
            return CheckResult("終了コード確認", "exit_code", NG, "タイムアウトにより異常終了")
        if actual == expected:
            return _apply_manual(
                CheckResult("終了コード確認", "exit_code", OK, f"終了コード = {actual}"),
                manual)
        return CheckResult("終了コード確認", "exit_code", NG, f"期待 {expected} に対し実績 {actual}")

    def _assert_db(self, spec: dict, result: CaseResult) -> List[CheckResult]:
        checks: List[CheckResult] = []
        for item in spec.get("db", []):
            table_name = item.get("table")
            actual = result.db_after.get(table_name)
            if actual is None:
                checks.append(
                    CheckResult(
                        f"DB照合: {table_name}", "db", NG,
                        f"実行後スナップショットに {table_name} がありません。snapshot.tables に定義してください。",
                    )
                )
                continue
            manual = item.get("manual") is True

            # 期待値なしで manual: true の場合、自動比較はせず証跡だけ残して人が見る
            if manual and not item.get("expected"):
                checks.append(CheckResult(
                    f"DB確認: {table_name}", "db", REVIEW,
                    "自動判定なし。実行前後のスナップショットを目視で確認してください。"))
                continue

            expected_path = self.settings.project_root / item["expected"]
            try:
                expected = db_mod.read_expected_csv(expected_path)
            except ConfigError as exc:
                checks.append(CheckResult(f"DB照合: {table_name}", "db", NG, str(exc)))
                continue
            result_check = compare.compare_db_table(
                table_name, actual, expected,
                keys=list(item.get("key") or []),
                ignore_columns=list(item.get("ignore_columns") or []),
            )
            checks.append(_apply_manual(result_check, manual))
        return checks

    def _assert_files(self, spec: dict) -> List[CheckResult]:
        checks: List[CheckResult] = []
        for item in spec.get("files", []):
            name = item.get("name") or item.get("expected") or "ファイル"

            if "exists" in item:
                checks.append(_apply_manual(
                    self._assert_file_exists(name, item["exists"]),
                    item.get("manual") is True))
                continue

            actual_spec = item.get("actual") or {}
            directory = self.settings.resolve_dir(actual_spec.get("dir", "output_dir"))
            pattern = self.settings.expand(actual_spec.get("pattern", "*"))
            matches = fsops.find_files(directory, pattern)
            if not matches:
                checks.append(
                    CheckResult(
                        f"ファイル照合: {name}", "file", NG,
                        f"{directory} に {pattern} が出力されていません",
                    )
                )
                continue
            # 複数一致は曖昧。先頭だけ比較すると、余分に出力された異常ファイルを
            # 見逃して OK になり得るため、既定では NG にする
            if len(matches) > 1:
                checks.append(
                    CheckResult(
                        f"ファイル照合: {name}", "file", NG,
                        f"パターン {pattern} に {len(matches)} 件が一致し曖昧です: "
                        f"{[m.name for m in matches[:5]]}（パターンを絞るか、余分な出力を調査してください）",
                    )
                )
                continue
            actual_path = matches[0]
            manual = item.get("manual") is True

            # 期待値なしで manual: true の場合、内容を証跡に残して人が判断する
            if manual and not item.get("expected"):
                checks.append(CheckResult(
                    f"ファイル確認: {name}", "file", REVIEW,
                    f"自動判定なし。出力内容を目視で確認してください: {actual_path.name}"))
                continue

            expected_path = self.settings.project_root / item["expected"]
            encoding = item.get("encoding", "utf-8")
            actual_text = fsops.read_text_file(actual_path, encoding, ["cp932", "utf-8"])
            checks.append(_apply_manual(
                compare.compare_text_file(
                    name, actual_path, expected_path,
                    actual_text=actual_text,
                    ignore_line_patterns=item.get("ignore_line_patterns"),
                ), manual))
        return checks

    def _assert_file_exists(self, name: str, exists_spec: dict) -> CheckResult:
        directory = self.settings.resolve_dir(exists_spec.get("dir", "output_dir"))
        pattern = self.settings.expand(exists_spec.get("pattern", "*"))
        expected_count = exists_spec.get("count")
        matches = fsops.find_files(directory, pattern)
        actual_count = len(matches)

        check_name = f"ファイル存在確認: {name}"
        detail = f"{directory} / {pattern} → {actual_count} 件"
        if expected_count is None:
            return CheckResult(check_name, "file", OK if matches else NG, detail)
        if actual_count == int(expected_count):
            return CheckResult(check_name, "file", OK, detail)
        return CheckResult(check_name, "file", NG, f"{detail}（期待 {expected_count} 件）")

    def _assert_log(self, spec: dict, log_slices: List[logs.LogSlice]) -> CheckResult:
        log_spec = spec.get("log")
        if not log_spec:
            return CheckResult("ログ確認", "log", SKIP, "期待値の指定なし")

        must = list(log_spec.get("must_contain") or [])
        must_not = list(log_spec.get("must_not_contain") or [])
        manual = log_spec.get("manual") is True
        if manual and not must and not must_not:
            return CheckResult(
                "ログ確認", "log", REVIEW,
                "自動判定なし。本実行で採取したログを目視で確認してください。")
        missing, forbidden = logs.check_keywords(log_slices, must, must_not)

        problems = []
        if missing:
            problems.append(f"必須キーワード未出力: {missing}")
        if forbidden:
            problems.append(f"禁止キーワード検出: {forbidden}")
        if problems:
            return CheckResult("ログ確認", "log", NG, " / ".join(problems))
        return _apply_manual(CheckResult(
            "ログ確認", "log", OK,
            f"必須 {len(must)} 件すべて検出 / 禁止 {len(must_not)} 件すべて未検出",
        ), manual)
