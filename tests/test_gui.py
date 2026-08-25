# -*- coding: utf-8 -*-
"""操作画面の子プロセス出力に関する回帰テスト。"""

import queue
import sys
import types
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

# CI や macOS の最小 Python には tkinter が無いことがある。ここで検査するのは
# GUI 部品ではなく subprocess の UTF-8 パイプなので、import 用の代用品で十分。
try:
    import tkinter  # noqa: F401
except ImportError:
    tkinter_stub = types.ModuleType("tkinter")
    tkinter_stub.messagebox = types.SimpleNamespace()
    tkinter_stub.ttk = types.SimpleNamespace()
    sys.modules["tkinter"] = tkinter_stub

from autotest import gui  # noqa: E402


class GuiSubprocessEncodingCase(unittest.TestCase):
    def test_utf8_output_does_not_use_windows_locale(self):
        """UTF-8 の日本語を cp932 locale に依存せず読み取れること。"""
        text = "設定確認：接続成功"
        proc = gui._popen(
            [sys.executable, "-c", "print(%r)" % text], ROOT)
        holder = types.SimpleNamespace(queue=queue.Queue())

        gui.AutotestGui._reader(holder, proc)

        self.assertEqual(holder.queue.get(), ("line", text + "\n"))
        self.assertEqual(holder.queue.get(), ("exit", 0))


if __name__ == "__main__":
    unittest.main()
