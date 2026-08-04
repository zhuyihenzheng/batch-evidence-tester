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
from .models import NG, OK, SKIP, CaseResult, CheckResult, ImageEvidence, Table
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

    # remove_dirs で消すフォルダへファイルを投入すると、投入時に再作成されて
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
            problems.append("投入ファイルがありません: %s" % src)
        dest_alias = spec.get("dest_dir", "input_dir")
        if dest_alias not in aliases:
            problems.append("input_files.dest_dir の論理名が paths に未定義です: %s" % dest_alias)

    # --- collect / folder_evidence ------------------------------------------
    for spec in (case.collect or {}).get("files", []):
        if spec.get("dir", "output_dir") not in aliases:
            problems.append("collect.files.dir の論理名が paths に未定義です: %s" % spec.get("dir"))
    for alias in (case.collect or {}).get("folder_evidence", []):
        if alias not in aliases:
            problems.append("collect.folder_evidence の論理名が paths に未定義です: %s" % alias)

    # --- 期待値ファイル -------------------------------------------------------
    for item in (case.assertions or {}).get("db", []):
        expected = settings.project_root / item.get("expected", "")
        if not item.get("expected"):
            problems.append("assert.db に expected がありません: %s" % item.get("table"))
        elif not expected.exists():
            problems.append("期待値 CSV がありません: %s" % expected)
    for item in (case.assertions or {}).get("files", []):
        if "expected" in item:
            expected = settings.project_root / item["expected"]
            if not expected.exists():
                problems.append("期待値ファイルがありません: %s" % expected)
        actual_spec = item.get("actual") or item.get("exists") or {}
        if actual_spec and actual_spec.get("dir", "output_dir") not in aliases:
            problems.append("assert.files の dir 論理名が paths に未定義です: %s" % actual_spec.get("dir"))

    return problems


class CaseRunner:
    def __init__(self, settings: Settings, run_dir: Path, offline: bool = False,
                 dry_run: bool = False, progress=None, default_date=None):
        self.settings = settings
        self._default_date = default_date
        self.run_dir = run_dir
        self.offline = offline
        self.dry_run = dry_run
        # 各工程の開始を通知するコールバック。処理が止まったとき、
        # どの工程で止まっているかを外から見えるようにするためのもの
        self._progress = progress or (lambda step: None)
        # batch 起動直前に DB から取る基準時刻（{batch_start} の展開に使う）
        self._batch_start_mark = None

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
        try:
            self._step("DB接続")
            client = db_mod.create_client(self.settings, self.offline, case.case_id, dry_run=self.dry_run)

            self._step("前処理（フォルダ準備・SQL・データ投入）")
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

            self._step("batch実行")
            result.execution = run_batch(
                self.settings,
                case.execute.get("args", []),
                dry_run=self.dry_run,
                batch_name=batch_name,
                on_progress=self._progress,   # 実行中の経過を run.log と画面へ
            )

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

            # 後処理（teardown）。証跡採取と判定が終わった後に実行する
            teardown_check = self._teardown(case, client)
            if teardown_check is not None:
                result.checks.append(teardown_check)

        except Exception as exc:  # 1 ケースの失敗で全体を止めない
            result.fatal_error = f"{type(exc).__name__}: {exc}\n{traceback.format_exc(limit=3)}"
        finally:
            if client:
                client.close()

        return result

    # ------------------------------------------------------------------
    def _setup(self, case: TestCase, client: db_mod.DbClient) -> None:
        setup = case.setup or {}

        # remove_dirs で「フォルダが存在しない」状態を作る異常系があるため、
        # そのフォルダは自動作成の対象から外す（作った直後に消すのは無駄で紛らわしい）
        removing = set(setup.get("remove_dirs", []))
        fsops.ensure_dirs(self.settings, [a for a in self.settings.path_aliases if a not in removing])

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
            table = client.snapshot(resolved, self.settings.db_format)
            target[table.title] = table

    def _resolve_batch_start(self, case: TestCase, client: db_mod.DbClient):
        """{batch_start} の基準時刻を決める。テスト機の時計を使う。

        DB へ問い合わせない分、権限やバージョン差の影響を受けず単純。
        ただしテスト機の時計が DB サーバより進んでいると、batch が最初に
        更新した行を取りこぼす（＝異常を見逃す）。これを避けるため
        既定で数秒さかのぼった時刻を基準にする。余分に拾う分には
        実行前スナップショットとの差で判別できるので害が小さい。
        """
        specs = case.snapshot.get("tables", [])
        if not any("{batch_start" in str(spec.get("sql", "")) for spec in specs):
            return None
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
        checks: List[CheckResult] = []

        checks.append(self._assert_exit_code(spec, result))
        checks.extend(self._assert_db(spec, result))
        checks.extend(self._assert_files(spec))
        checks.append(self._assert_log(spec, log_slices))

        return [c for c in checks if c is not None]

    def _assert_exit_code(self, spec: dict, result: CaseResult) -> CheckResult:
        if "exit_code" not in spec:
            return CheckResult("終了コード確認", "exit_code", SKIP, "期待値の指定なし")
        expected = spec["exit_code"]
        actual = result.execution.exit_code
        if result.execution.timed_out:
            return CheckResult("終了コード確認", "exit_code", NG, "タイムアウトにより異常終了")
        if actual == expected:
            return CheckResult("終了コード確認", "exit_code", OK, f"終了コード = {actual}")
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
            expected_path = self.settings.project_root / item["expected"]
            try:
                expected = db_mod.read_expected_csv(expected_path)
            except ConfigError as exc:
                checks.append(CheckResult(f"DB照合: {table_name}", "db", NG, str(exc)))
                continue
            checks.append(
                compare.compare_db_table(
                    table_name, actual, expected,
                    keys=list(item.get("key") or []),
                    ignore_columns=list(item.get("ignore_columns") or []),
                )
            )
        return checks

    def _assert_files(self, spec: dict) -> List[CheckResult]:
        checks: List[CheckResult] = []
        for item in spec.get("files", []):
            name = item.get("name") or item.get("expected") or "ファイル"

            if "exists" in item:
                checks.append(self._assert_file_exists(name, item["exists"]))
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
            expected_path = self.settings.project_root / item["expected"]
            encoding = item.get("encoding", "utf-8")
            actual_text = fsops.read_text_file(actual_path, encoding, ["cp932", "utf-8"])
            checks.append(
                compare.compare_text_file(
                    name, actual_path, expected_path,
                    actual_text=actual_text,
                    ignore_line_patterns=item.get("ignore_line_patterns"),
                )
            )
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
        missing, forbidden = logs.check_keywords(log_slices, must, must_not)

        problems = []
        if missing:
            problems.append(f"必須キーワード未出力: {missing}")
        if forbidden:
            problems.append(f"禁止キーワード検出: {forbidden}")
        if problems:
            return CheckResult("ログ確認", "log", NG, " / ".join(problems))
        return CheckResult(
            "ログ確認", "log", OK,
            f"必須 {len(must)} 件すべて検出 / 禁止 {len(must_not)} 件すべて未検出",
        )
