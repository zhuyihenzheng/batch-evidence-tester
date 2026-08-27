"""操作画面（tkinter）。

テスト実施者が「一覧を見て、選んで、実行して、証跡を開く」までを
コマンドを覚えずに行えるようにするためのもの。

設計方針:
  - tkinter しか使わない。閉じた社内ネットワークでは追加インストールが
    できず、ブラウザも何が入っているか当てにできない。Tk は Anaconda に同梱。
  - 実行は必ず subprocess で CLI を起動する。GUI 内で CaseRunner を直接
    呼ばない。理由は 3 つ:
      1. 数十分かかる実行や pyodbc の異常終了で画面を巻き添えにしない
      2. tkinter はスレッドセーフではない
      3. 編排の入口を CLI 1 本に保つ（画面用の 2 本目を作ると必ずずれる）
  - 進捗は子プロセスの標準出力を読んで表示する。読み取りはデーモンスレッド、
    画面更新は after() で行う（Tk のウィジェットを触るのはメインスレッドだけ）。
  - 無人実行は従来どおり CLI + タスクスケジューラ。この画面は対話用。

  python -m autotest gui
"""

import os
import subprocess
import sys
import threading
from pathlib import Path
from typing import List, Optional

try:
    import queue
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError as exc:  # pragma: no cover - GUI 無し環境
    raise SystemExit(
        "画面を開けません（tkinter が見つかりません）: %s\n"
        "  コマンドラインからは通常どおり使えます: python -m autotest run" % exc)

from .config import ConfigError, load_cases, load_settings

CHECKED = "☑"
UNCHECKED = "☐"

# 終了コードの意味。run_test.bat の表示と揃える
VERDICT_TEXT = {
    0: ("総合判定: OK", "#1a7f37"),
    1: ("総合判定: NG  ―  Excel の「サマリ」シートを確認してください", "#b42318"),
    2: ("設定エラーで実行できませんでした", "#b42318"),
    3: ("総合判定: 要確認  ―  Excel の黄色い欄を記入後、finalize してください", "#9a6700"),
}


def _python_executable() -> str:
    """子プロセスを起動する python。

    pythonw.exe で GUI を起動している場合、そのまま sys.executable を使うと
    子プロセスもコンソール無しになる。標準出力を PIPE で受けるので実害は
    無いが、python.exe があればそちらを使う方が挙動が素直になる。
    """
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        console = exe.with_name("python.exe")
        if console.exists():
            return str(console)
    return str(exe)


def _popen(args: List[str], project_root: Path) -> subprocess.Popen:
    """CLI を子プロセスとして起動する。"""
    env = dict(os.environ)
    env["PYTHONPATH"] = str(project_root / "src")
    # cp932 では出せない記号があると子プロセス側が UnicodeEncodeError で落ちる。
    # 証跡には影響しないので置換して流す
    env["PYTHONIOENCODING"] = "utf-8:replace"
    env["PYTHONUNBUFFERED"] = "1"

    kwargs = {}
    if os.name == "nt":
        # pythonw から起動したとき黒い窓を出さない
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

    return subprocess.Popen(
        args,
        cwd=str(project_root),
        env=env,
        # pythonw ではコンソールが無い。明示しないと無効なハンドルを渡してしまう
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        # バイナリパイプは行バッファリング非対応。子側は PYTHONUNBUFFERED=1。
        bufsize=0,
        **kwargs
    )


def _open_with_os(path: Path) -> None:
    """既定のアプリケーションで開く（Excel / フォルダ）。"""
    if os.name == "nt":
        os.startfile(str(path))  # noqa: S606 - Windows の標準的な開き方
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class AutotestGui(object):
    """メイン画面。"""

    def __init__(self, root, project_root: Path, config_path: Optional[Path] = None,
                 cases_dir: Optional[Path] = None, offline: bool = False):
        self.root = root
        self.project_root = project_root
        self.config_path = config_path
        self.cases_dir = cases_dir or (project_root / "cases")
        self.cases = []
        self.checked = set()
        self.proc = None
        self.queue = queue.Queue()
        self.last_excel = None
        self.last_log = None
        self.settings = None
        self.current_kind = ""
        # DB 接続の確認状態。既定は「未確認」——確認していないものを
        # 大丈夫そうに見せない。接続先を取り違えたまま試験すると、
        # 結果は正常に見えるのに別の DB を見ている、という最悪の偽 OK になる
        self.db_state = "unknown"
        self.db_actual = ""
        self._db_actual = {}
        self._db_mismatch = False

        root.title("バッチ自動テスト")
        root.geometry("1180x760")
        self._build_widgets()
        self.refresh_cases()
        self.root.after(200, self._poll)

    # ------------------------------------------------------------------
    def _build_widgets(self) -> None:
        top = ttk.Frame(self.root, padding=(10, 8))
        top.pack(fill="x")

        ttk.Label(top, text="設定ファイル:").pack(side="left")
        self.config_label = ttk.Label(top, text="-", foreground="#444")
        self.config_label.pack(side="left", padx=(4, 16))

        ttk.Label(top, text="タグ:").pack(side="left")
        self.tag_var = tk.StringVar(value="（すべて）")
        self.tag_box = ttk.Combobox(top, textvariable=self.tag_var, state="readonly", width=16)
        self.tag_box.pack(side="left", padx=4)
        self.tag_box.bind("<<ComboboxSelected>>", lambda e: self._render_rows())

        ttk.Button(top, text="再読込", command=self.refresh_cases).pack(side="left", padx=4)
        self.offline_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="オフライン（DB を使わない）",
                        variable=self.offline_var,
                        command=self._render_db_bar).pack(side="left", padx=12)

        # --- 接続先 DB ------------------------------------------------------
        # 常時表示する。「どこに繋ぐか」が見えない状態で実行させないため
        dbbar = tk.Frame(self.root, padx=10, pady=6)
        dbbar.pack(fill="x")
        self.db_frame = tk.Frame(dbbar, padx=10, pady=6, relief="solid", borderwidth=1)
        self.db_frame.pack(fill="x")

        # 右側から先に pack する。pack は先に詰めたものが幅を確保するので、
        # 逆順にすると接続先の文字列が長いときに状態表示が切れて読めなくなる
        self.db_check_button = ttk.Button(self.db_frame, text="DB接続確認",
                                          command=self.dbcheck)
        self.db_check_button.pack(side="right")
        self.db_status = tk.Label(self.db_frame, text="", anchor="e", font=("", 12, "bold"))
        self.db_status.pack(side="right", padx=(0, 12))
        self.db_env = tk.Label(self.db_frame, text="", anchor="w", font=("", 12, "bold"),
                               fg="#222")
        self.db_env.pack(side="left")
        self.db_target = tk.Label(self.db_frame, text="", anchor="w")
        self.db_target.pack(side="left", padx=(12, 0))

        # --- ケース一覧 -----------------------------------------------------
        middle = ttk.Frame(self.root, padding=(10, 0))
        middle.pack(fill="both", expand=True)

        columns = ("check", "case_id", "tags", "mode", "name")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", height=12)
        for key, title, width, anchor in (
            ("check", "選択", 44, "center"),
            ("case_id", "ケースID", 200, "w"),
            ("tags", "区分", 150, "w"),
            ("mode", "実行方式", 80, "center"),
            ("name", "ケース名", 460, "w"),
        ):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor)
        self.tree.tag_configure("manual", foreground="#9a6700")

        scroll = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="left", fill="y")
        self.tree.bind("<Button-1>", self._on_click)

        # --- 操作ボタン -----------------------------------------------------
        bar = ttk.Frame(self.root, padding=(10, 8))
        bar.pack(fill="x")
        self.buttons = []
        for text, command in (
            ("全件実行", self.run_all),
            ("選択したケースを実行", self.run_selected),
            ("タグで実行", self.run_by_tag),
            ("設定確認 (validate)", self.validate),
        ):
            b = ttk.Button(bar, text=text, command=command)
            b.pack(side="left", padx=3)
            self.buttons.append(b)

        bar2 = ttk.Frame(self.root, padding=(10, 0))
        bar2.pack(fill="x")
        for text, command in (
            ("新規ケース作成", self.new_case),
            ("選択ケースを複製", self.copy_case),
            ("手動：実行前を採取", self.manual_before),
            ("手動：実行後を採取", self.manual_after),
        ):
            b = ttk.Button(bar2, text=text, command=command)
            b.pack(side="left", padx=3)
            self.buttons.append(b)

        self.cancel_button = ttk.Button(bar2, text="中止", command=self.cancel_run,
                                        state="disabled")
        self.cancel_button.pack(side="left", padx=12)
        self.excel_button = ttk.Button(bar2, text="証跡Excelを開く", command=self.open_excel,
                                       state="disabled")
        self.excel_button.pack(side="left", padx=3)

        # --- 実行ログ -------------------------------------------------------
        bottom = ttk.Frame(self.root, padding=(10, 8))
        bottom.pack(fill="both", expand=True)
        self.status = tk.Label(bottom, text="待機中", anchor="w", font=("", 11, "bold"))
        self.status.pack(fill="x")
        self.log = tk.Text(bottom, height=12, wrap="none", state="disabled",
                           background="#1e1e1e", foreground="#e6e6e6")
        log_scroll = ttk.Scrollbar(bottom, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="left", fill="y")

    # ------------------------------------------------------------------
    def refresh_cases(self) -> None:
        """ケース定義を読み直す。読めない場合は理由をそのまま見せる。"""
        try:
            settings_path = self.config_path or _default_config(self.project_root)
            settings = load_settings(settings_path, project_root=self.project_root)
            if self.settings is not None and self._db_signature(settings) != self._db_signature(self.settings):
                # 設定を差し替えたら、前の確認結果は当てにならない
                self.db_state = "unknown"
                self.db_actual = ""
            self.settings = settings
            self.config_label.config(text=str(settings_path.name))
            # enabled: false も含めて見せたいが、load_cases は除外する。
            # 実行できるものだけを一覧に出す方が「押したのに動かない」を防げる
            self.cases = load_cases(self.cases_dir)
        except ConfigError as exc:
            self.cases = []
            messagebox.showerror("設定エラー", str(exc))
        self._render_db_bar()

        tags = sorted({t for c in self.cases for t in c.tags})
        self.tag_box["values"] = ["（すべて）"] + tags
        if self.tag_var.get() not in self.tag_box["values"]:
            self.tag_var.set("（すべて）")
        # 初回表示と「再読込」の後は未選択から始める。再読込は最新の定義を
        # 読み直す操作であり、全ケースを暗黙に実行対象へ加えない。
        self.checked = set()
        self._render_rows()

    @staticmethod
    def _db_signature(settings) -> tuple:
        """接続先が変わったかを見るための識別子。"""
        db = settings.database
        return (str(db.get("server", "")), str(db.get("database", "")),
                str(db.get("auth", "")), str(db.get("user", "")))

    def _render_db_bar(self) -> None:
        """接続先 DB と、その確認状態を表示する。

        「確認していない」ことをはっきり出す。接続先を取り違えたまま試験すると
        結果は正常に見えるのに別の DB を見ている状態になり、証跡としては
        最悪の部類の誤りになる。だから既定は緑ではなく灰色の「未確認」。
        """
        if self.settings is None:
            self.db_env.config(text="")
            self.db_target.config(text="設定を読み込めていません", fg="#b42318")
            self.db_status.config(text="")
            self.db_check_button.config(state="disabled")
            return

        env_name = str(self.settings.env.get("name", "")) or "（環境名なし）"
        db = self.settings.database
        auth = "Windows認証" if str(db.get("auth", "sql")).lower() == "windows" \
            else "SQL Server認証: %s" % db.get("user", "")
        self.db_env.config(text="環境: %s" % env_name)
        self.db_target.config(
            text="接続先: %s / %s   （%s）" % (db.get("server", ""), db.get("database", ""), auth),
            fg="#333")

        if self.offline_var.get():
            # offline は fixtures/ の CSV を使うので DB へは一切繋がない
            self._paint_db_bar("オフライン（DB は使いません）", "#555", "#eeeeee")
            self.db_check_button.config(state="disabled")
            return
        self.db_check_button.config(state="normal" if self.proc is None else "disabled")

        text, fg, bg = {
            "unknown": ("● 未確認（実行前に確認）", "#9a6700", "#fff8e1"),
            "ok": ("● 確認済: %s" % self.db_actual, "#1a7f37", "#eaf7ee"),
            "mismatch": ("▲ 接続先が設定と相違: %s" % self.db_actual, "#b42318", "#fdeceb"),
            "ng": ("✕ 接続できません", "#b42318", "#fdeceb"),
        }[self.db_state]
        self._paint_db_bar(text, fg, bg)

    def _paint_db_bar(self, status_text: str, status_fg: str, bg: str) -> None:
        """枠と中のラベルをまとめて塗る。tk.Label は親の背景色を継承しない。"""
        self.db_frame.config(background=bg)
        for label in (self.db_env, self.db_target):
            label.config(background=bg)
        self.db_status.config(text=status_text, fg=status_fg, background=bg)

    def dbcheck(self) -> None:
        """SQL Server へ実際に接続して、繋がる先を確認する。"""
        self._start(self._base_args() + ["dbcheck"] + self._config_only(), "DB接続確認",
                    kind="dbcheck")

    def _visible_cases(self) -> list:
        tag = self.tag_var.get()
        if tag and tag != "（すべて）":
            return [c for c in self.cases if tag in c.tags]
        return list(self.cases)

    def _render_rows(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for case in self._visible_cases():
            mark = CHECKED if case.case_id in self.checked else UNCHECKED
            mode = "手動" if case.is_manual else "自動"
            self.tree.insert(
                "", "end", iid=case.case_id,
                values=(mark, case.case_id, "/".join(case.tags), mode, case.name),
                tags=("manual",) if case.is_manual else ())
        self.status.config(text="%d 件表示中（%d 件が定義済み）"
                                % (len(self._visible_cases()), len(self.cases)))

    def _on_click(self, event) -> None:
        """選択列のクリックでチェックを切り替える。"""
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        case_id = self.tree.identify_row(event.y)
        if not case_id:
            return
        if case_id in self.checked:
            self.checked.discard(case_id)
        else:
            self.checked.add(case_id)
        self._render_rows()

    def _selected_case(self):
        """1 件だけ選ばれているケース（複製・手動採取で使う）。"""
        focused = self.tree.focus()
        for case in self.cases:
            if case.case_id == focused:
                return case
        return None

    # ------------------------------------------------------------------
    def _base_args(self) -> List[str]:
        args = [_python_executable(), "-m", "autotest"]
        return args

    def _config_only(self) -> List[str]:
        """--config だけ。dbcheck は offline と無関係（実接続を確認するコマンド）。"""
        return ["--config", str(self.config_path)] if self.config_path else []

    def _common_options(self) -> List[str]:
        options = self._config_only()
        # GUI 下部の実行ログにも CaseRunner の工程（前置 batch のコマンド、
        # stdout、stderr を含む）を逐次表示する。
        options.append("--verbose")
        if self.offline_var.get():
            options.append("--offline")
        return options

    def _start(self, command: List[str], label: str, kind: str = "run") -> None:
        if self.proc is not None:
            messagebox.showinfo("実行中", "実行中です。終わるまでお待ちください。")
            return

        self._clear_log()
        self._append_log("$ " + " ".join(command[2:]) + "\n")
        self.status.config(text=label + " 実行中...", foreground="#0b5cad")
        self.current_kind = kind
        self.last_excel = None
        self.excel_button.config(state="disabled")
        for b in self.buttons:
            b.config(state="disabled")
        self.db_check_button.config(state="disabled")
        self.cancel_button.config(state="normal")
        if kind == "dbcheck":
            self.db_actual = ""
            self._db_actual = {}
            self._db_mismatch = False

        try:
            self.proc = _popen(command, self.project_root)
        except OSError as exc:
            self._finish_run(None, "起動できませんでした: %s" % exc)
            return

        thread = threading.Thread(target=self._reader, args=(self.proc,))
        thread.daemon = True   # 画面を閉じたときに残らないように
        thread.start()

    def _reader(self, proc) -> None:
        """子プロセスの出力を読む。ウィジェットには触らない（別スレッドのため）。"""
        try:
            # 子プロセスには PYTHONIOENCODING=utf-8 を指定しているため、ここでも
            # UTF-8 として明示的に読む。universal_newlines=True に任せると日本語
            # Windows では locale の cp932 が選ばれ、UTF-8 の日本語を読んだ時点で
            # UnicodeDecodeError になって進捗読取スレッドが停止してしまう。
            for raw_line in iter(proc.stdout.readline, b""):
                line = (raw_line.decode("utf-8", "replace")
                        if not isinstance(raw_line, str) else raw_line)
                self.queue.put(("line", line))
        finally:
            proc.stdout.close()
            self.queue.put(("exit", proc.wait()))

    def _poll(self) -> None:
        """キューを吸い出して画面へ反映する。メインスレッドで動く。"""
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "line":
                    self._append_log(payload)
                    self._scan_line(payload)
                else:
                    self._finish_run(payload)
        except queue.Empty:
            pass
        self.root.after(200, self._poll)

    def _scan_line(self, line: str) -> None:
        """CLI の出力から成果物のパスと、DB の実接続先を拾う。"""
        for prefix, attr in (("証跡Excel:", "last_excel"), ("実行ログ :", "last_log")):
            if line.startswith(prefix):
                setattr(self, attr, Path(line[len(prefix):].strip()))

        if self.current_kind != "dbcheck":
            return
        stripped = line.strip()
        # dbcheck は「設定値」ではなく DB に問い合わせた実際の接続先を出す。
        # 表示すべきはこちら —— 設定を読み上げるだけでは取り違えを検出できない
        for prefix, key in (("サーバ名   :", "server"), ("データベース:", "database")):
            if stripped.startswith(prefix):
                self._db_actual[key] = stripped[len(prefix):].strip()
        # 成否そのものは終了コードで判断する（出力の途中にも [NG] は出るため）。
        # ここで拾うのは「繋がったが設定と違う先だった」という、
        # 終了コードだけでは区別できない事実
        if stripped.startswith("[警告] 接続先 DB が設定値"):
            self._db_mismatch = True

    def _finish_run(self, code: Optional[int], error: str = "") -> None:
        self.proc = None
        for b in self.buttons:
            b.config(state="normal")
        self.cancel_button.config(state="disabled")

        if error:
            self.status.config(text=error, foreground="#b42318")
            self._render_db_bar()
            return

        if self.current_kind == "dbcheck":
            # 成否は終了コードで決める。繋がったうえで設定と違う先だった場合だけ
            # 出力から拾った mismatch を優先する（この場合 dbcheck 自体も 1 を返す）
            if code == 0 and not self._db_mismatch:
                self.db_state = "ok"
            elif self._db_mismatch:
                self.db_state = "mismatch"
            else:
                self.db_state = "ng"
            self.db_actual = "%s / %s" % (self._db_actual.get("server", "?"),
                                          self._db_actual.get("database", "?"))
            self.refresh_cases()
            self.status.config(
                text={"ok": "DB接続確認: OK", "mismatch": "DB接続確認: 接続先が設定と違います",
                      "ng": "DB接続確認: 繋がりません"}[self.db_state],
                foreground="#1a7f37" if self.db_state == "ok" else "#b42318")
            if self.db_state == "mismatch":
                messagebox.showwarning(
                    "接続先の相違",
                    "設定した DB と、実際に繋がった DB が違います。\n\n"
                    "  実際の接続先: %s\n\n"
                    "このまま試験すると、別の DB を見た結果が証跡として残ります。\n"
                    "settings の database.database と、ログインの既定データベースを"
                    "確認してください。" % self.db_actual)
            return

        # 一覧の更新を先に行う。refresh_cases は件数をステータスへ書くので、
        # 後に呼ぶと判定結果の表示を上書きしてしまう
        # （new / copy の直後は一覧が変わっているので更新は必要）
        self.refresh_cases()
        if self.last_excel and self.last_excel.exists():
            self.excel_button.config(state="normal")
        text, color = VERDICT_TEXT.get(code, ("終了コード %s" % code, "#444"))
        self.status.config(text=text, foreground=color)

    # --- 実行系 ---------------------------------------------------------
    def run_all(self) -> None:
        self._start(self._base_args() + ["run"] + self._common_options(), "全件")

    def run_selected(self) -> None:
        visible = {c.case_id for c in self._visible_cases()}
        chosen = [c for c in self.cases if c.case_id in self.checked and c.case_id in visible]
        if not chosen:
            messagebox.showinfo("選択なし", "実行するケースを選んでください。")
            return
        manual = [c.case_id for c in chosen if c.is_manual]
        if manual:
            messagebox.showinfo(
                "手動実施ケース",
                "次のケースは手動実施です。自動実行の選択から外してください:\n  %s\n\n"
                "採取は「手動：実行前を採取」から行います。" % ", ".join(manual))
            return
        args = []
        for case in chosen:
            args += ["--case", case.case_id]
        self._start(self._base_args() + ["run"] + args + self._common_options(),
                    "%d 件" % len(chosen))

    def run_by_tag(self) -> None:
        tag = self.tag_var.get()
        if not tag or tag == "（すべて）":
            messagebox.showinfo("タグ未指定", "上のタグ欄で実行したいタグを選んでください。")
            return
        self._start(self._base_args() + ["run", "--tag", tag] + self._common_options(),
                    "タグ %s" % tag)

    def validate(self) -> None:
        self._start(self._base_args() + ["validate"] + self._common_options(), "設定確認")

    def cancel_run(self) -> None:
        if self.proc is None:
            return
        if not messagebox.askyesno(
                "中止の確認",
                "実行を中止します。\n\n"
                "途中で止めるため証跡は不完全になります。\n"
                "また、テスト対象のフォルダや DB は途中の状態のまま残ります。\n\n"
                "中止してよろしいですか？"):
            return
        self._append_log("\n--- 中止しました ---\n")
        _kill_tree(self.proc)

    def open_excel(self) -> None:
        if self.last_excel and self.last_excel.exists():
            _open_with_os(self.last_excel)

    # --- ケース作成 -----------------------------------------------------
    def new_case(self) -> None:
        from . import scaffold

        dialog = _NewCaseDialog(self.root, scaffold.list_templates(self.project_root))
        if not dialog.result:
            return
        case_id, template = dialog.result
        self._start(self._base_args() + ["new", "--id", case_id, "--template", template]
                    + self._common_options(), "ケース作成")

    def copy_case(self) -> None:
        case = self._selected_case()
        if case is None:
            messagebox.showinfo("選択なし", "複製したいケースを一覧でクリックして選んでください。")
            return
        dialog = _NewCaseDialog(self.root, None, title="ケースの複製",
                                note="複製元: %s" % case.case_id)
        if not dialog.result:
            return
        new_id = dialog.result[0]
        self._start(self._base_args() + ["copy", "--from", case.case_id, "--id", new_id]
                    + self._common_options(), "ケース複製")

    # --- 手動実施 -------------------------------------------------------
    def _manual(self, phase: str) -> None:
        case = self._selected_case()
        if case is None:
            messagebox.showinfo("選択なし", "手動実施するケースを一覧でクリックして選んでください。")
            return
        if not case.is_manual:
            messagebox.showinfo(
                "自動実行ケース",
                "%s は自動実行のケース（mode: auto）です。\n"
                "「選択したケースを実行」から実行してください。" % case.case_id)
            return
        self._start(self._base_args() + ["manual", "--case", case.case_id, "--phase", phase]
                    + self._common_options(), "手動採取（%s）" % phase)

    def manual_before(self) -> None:
        self._manual("before")

    def manual_after(self) -> None:
        self._manual("after")

    # --- ログ表示 -------------------------------------------------------
    def _append_log(self, text: str) -> None:
        self.log.config(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear_log(self) -> None:
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")


def _kill_tree(proc) -> None:
    """子プロセスを孫ごと止める。batch(.exe) が残ると後続が動かせない。"""
    if os.name == "nt":
        subprocess.call(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        proc.terminate()


class _NewCaseDialog(object):
    """ケースIDとひな形を尋ねる小さなダイアログ。"""

    def __init__(self, parent, templates: Optional[List[str]],
                 title: str = "新しいケースの作成", note: str = ""):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.resizable(False, False)

        frame = ttk.Frame(self.top, padding=14)
        frame.pack(fill="both", expand=True)

        if note:
            ttk.Label(frame, text=note, foreground="#444").grid(
                row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        ttk.Label(frame, text="ケースID:").grid(row=1, column=0, sticky="w")
        self.id_var = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=self.id_var, width=32)
        entry.grid(row=1, column=1, sticky="we", padx=(8, 0))
        entry.focus_set()

        self.template_var = tk.StringVar(value=(templates or [""])[0])
        if templates:
            ttk.Label(frame, text="ひな形:").grid(row=2, column=0, sticky="w", pady=(8, 0))
            ttk.Combobox(frame, textvariable=self.template_var, values=templates,
                         state="readonly", width=30).grid(
                row=2, column=1, sticky="we", padx=(8, 0), pady=(8, 0))

        ttk.Label(frame, text="※ 作成したケースは enabled: false です。\n"
                              "   投入ファイルを置いて定義を整えてから true にしてください。",
                  foreground="#666").grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 0))

        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="キャンセル", command=self.top.destroy).pack(side="right", padx=4)
        ttk.Button(buttons, text="作成", command=self._ok).pack(side="right")
        self.top.bind("<Return>", lambda e: self._ok())

        parent.wait_window(self.top)

    def _ok(self) -> None:
        case_id = self.id_var.get().strip()
        if not case_id:
            messagebox.showwarning("入力なし", "ケースIDを入力してください。", parent=self.top)
            return
        self.result = (case_id, self.template_var.get())
        self.top.destroy()


def _default_config(project_root: Path) -> Path:
    local = project_root / "config" / "settings.local.yaml"
    return local if local.exists() else project_root / "config" / "settings.yaml"


def main(project_root: Path, config_path: Optional[Path] = None,
         cases_dir: Optional[Path] = None) -> int:
    root = tk.Tk()
    AutotestGui(root, project_root, config_path=config_path, cases_dir=cases_dir)
    root.mainloop()
    return 0
