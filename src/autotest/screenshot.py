"""実画面キャプチャ（evidence.mode = screen / both のとき使用）。

「Explorer を実際に開いた画面が欲しい」という要求に応えるための任意機能。
対話ログオンセッションが必須で、無人実行では使えない。失敗しても実行は止めず、
render 方式の画像へフォールバックする。
"""

import platform
import subprocess
import time
from pathlib import Path
from typing import Optional, Tuple


class ScreenCaptureUnavailable(Exception):
    """キャプチャ環境が整っていない。呼び出し側は render にフォールバックする。"""


def is_available() -> Tuple[bool, str]:
    if platform.system() != "Windows":
        return False, "実画面キャプチャは Windows のみ対応です"
    try:
        import mss  # noqa: F401,PLC0415
    except ImportError:
        return False, "mss が未インストールです（pip install mss）"
    return True, ""


def capture_explorer(directory: Path, out_path: Path, wait_sec: float = 2.5) -> Path:
    """Explorer でフォルダを開き、前面ウィンドウをキャプチャする。"""
    ok, reason = is_available()
    if not ok:
        raise ScreenCaptureUnavailable(reason)

    subprocess.Popen(["explorer", str(directory)])  # noqa: S603,S607
    time.sleep(wait_sec)
    return capture_foreground(out_path)


def capture_foreground(out_path: Path) -> Path:
    """前面ウィンドウ（取得できなければプライマリモニタ全体）をキャプチャする。"""
    ok, reason = is_available()
    if not ok:
        raise ScreenCaptureUnavailable(reason)

    import mss  # noqa: PLC0415
    from PIL import Image  # noqa: PLC0415

    region = _foreground_window_rect()
    with mss.mss() as sct:
        monitor = region or sct.monitors[1]
        shot = sct.grab(monitor)
        img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path)
    return out_path


def _foreground_window_rect() -> Optional[dict]:
    """pywin32 があれば前面ウィンドウの矩形を返す。無ければ None（全画面にフォールバック）。"""
    try:
        import win32gui  # noqa: PLC0415
    except ImportError:
        return None
    try:
        hwnd = win32gui.GetForegroundWindow()
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        if right - left < 100 or bottom - top < 100:
            return None
        return {"left": left, "top": top, "width": right - left, "height": bottom - top}
    except Exception:
        return None
