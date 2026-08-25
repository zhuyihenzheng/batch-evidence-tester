# -*- coding: utf-8 -*-
"""実行環境が AUTO_TEST_BATCH の要件を満たすか確認する。

このスクリプト自体は Python 2.7 / 3.6 でも動く構文で書いてある。
本体が動かない環境でも「なぜ動かないか」を出せるようにするため、
f-string・型ヒント・dataclass の類は一切使っていない。

  python check_env.py
"""

import os
import platform
import subprocess
import sys

# Anaconda 5.2 (Python 3.6.5) で動くよう本体を書いてあるため、下限は 3.6
REQUIRED_PYTHON = (3, 6)
RECOMMENDED_PYTHON = (3, 6)

# (import 名, 表示名, 最低バージョン, 理由)
# Anaconda 5.2 同梱版（openpyxl 2.5.3 / Pillow 5.1.0 / PyYAML 3.12）で動く下限にしてある
REQUIRED_PACKAGES = [
    ("openpyxl", "openpyxl", (2, 5), "Excel 成果物の生成"),
    ("PIL", "Pillow", (5, 1), "証跡画像の描画"),
    ("yaml", "PyYAML", (3, 12), "設定・ケース定義の読み込み"),
]
OPTIONAL_PACKAGES = [
    ("pyodbc", "pyodbc", (4, 0), "SQL Server 接続（--offline のみで使うなら不要）"),
    ("mss", "mss", (6, 0), "実画面キャプチャ（evidence.mode = screen / both のときのみ）"),
]


def parse_version(text):
    """'3.1.5' -> (3, 1, 5)。数字以外の要素が来たら 0 に丸める。"""
    parts = []
    for chunk in str(text).split("."):
        digits = ""
        for ch in chunk:
            if ch.isdigit():
                digits += ch
            else:
                break
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def _run(command):
    """外部コマンドを実行して標準出力を返す。失敗したら None。

    subprocess.run(capture_output=) は Python 3.7 以降のため、
    古い環境でも動くよう Popen で書いている。
    """
    try:
        proc = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=False)
        out, _err = proc.communicate()
    except Exception:
        return None
    if proc.returncode != 0:
        return None
    if not isinstance(out, str):
        out = out.decode("utf-8", "replace")
    return out.strip()


def report_conda():
    """Anaconda / Miniconda 環境かどうかと、そのバージョンを表示する。

    「Anaconda のバージョン」は conda 自身のバージョンとは別物なので、
    ディストリビューション版（anaconda メタパッケージ）も併せて出す。
    """
    is_conda = ("conda" in sys.version.lower() or "anaconda" in sys.version.lower()
                or os.path.isdir(os.path.join(sys.prefix, "conda-meta")))
    if not is_conda:
        print("conda       : 使用していません（通常の Python）")
        return

    env_name = os.environ.get("CONDA_DEFAULT_ENV", "(不明)")
    print("conda 環境  : %s   prefix=%s" % (env_name, sys.prefix))

    conda_version = _run(["conda", "--version"])
    print("conda       : %s" % (conda_version or "コマンドが見つかりません（PATH 未設定）"))

    listed = _run(["conda", "list", "anaconda"])
    distro = None
    if listed:
        for line in listed.splitlines():
            parts = line.split()
            # 例: "anaconda   5.2.0   py36_3"   ※ anaconda-client 等は除外する
            if len(parts) >= 2 and parts[0] == "anaconda":
                distro = parts[1]
                break
    if distro:
        print("Anaconda 版 : %s" % distro)
        if parse_version(distro) < (5, 3):
            print("              ※ Python 3.6 系ですが、本体は 3.6 対応済みのため動作します")
    else:
        print("Anaconda 版 : メタパッケージなし（Miniconda か、個別構築の環境）")


def check_python():
    actual = sys.version_info[:3]
    print("Python      : %d.%d.%d  (%s)" % (actual[0], actual[1], actual[2], sys.executable))
    print("OS          : %s %s" % (platform.system(), platform.release()))
    report_conda()

    if actual[:2] < REQUIRED_PYTHON:
        print("")
        print("  [NG] Python %d.%d 以上が必要です。" % REQUIRED_PYTHON)
        print("       Python 2 系、または 3.5 以前では動作しません。")
        return False
    return True


def check_package(import_name, display_name, min_version, reason, required):
    label = "必須" if required else "任意"
    try:
        module = __import__(import_name)
    except ImportError:
        mark = "NG" if required else "--"
        print("  [%s] %-10s 未インストール       (%s: %s)" % (mark, display_name, label, reason))
        return not required

    version = getattr(module, "__version__", None) or getattr(module, "VERSION", None) or "?"
    if version == "?":
        print("  [??] %-10s バージョン不明        (%s: %s)" % (display_name, label, reason))
        return True

    if parse_version(version) < min_version:
        need = ".".join(str(n) for n in min_version)
        mark = "NG" if required else "--"
        print("  [%s] %-10s %-12s → %s 以上が必要 (%s)" % (mark, display_name, version, need, reason))
        return not required

    print("  [OK] %-10s %s" % (display_name, version))
    return True


def main():
    print("=" * 74)
    print(" AUTO_TEST_BATCH  実行環境チェック")
    print("=" * 74)

    ok = check_python()

    print("")
    print("必須パッケージ:")
    for import_name, display_name, min_version, reason in REQUIRED_PACKAGES:
        if not check_package(import_name, display_name, min_version, reason, True):
            ok = False

    print("")
    print("任意パッケージ:")
    for import_name, display_name, min_version, reason in OPTIONAL_PACKAGES:
        check_package(import_name, display_name, min_version, reason, False)

    # 操作画面（run_gui.bat）用。標準ライブラリなので通常は入っているが、
    # Linux の一部ディストリや最小構成の Python では別パッケージになっている。
    # 無くてもコマンドラインからは全機能が使えるため、任意扱いにする。
    try:
        import tkinter
        print("  [OK] tkinter    Tk %s  操作画面 run_gui.bat 用" % tkinter.TkVersion)
    except ImportError:
        print("  [--] tkinter    未導入  → 操作画面は使えません")
        print("                  （コマンドラインからは通常どおり全機能が使えます）")

    print("")
    print("=" * 74)
    if ok:
        print(" 結論: この環境で実行できます。")
        print("   必須パッケージは導入済みです。conda / pip の追加実行は不要です。")
        print("")
        if os.name == "nt":
            print("   次は、プロジェクト直下の Command Prompt で以下を実行してください:")
            print("     set PYTHONPATH=%CD%\\src")
        else:
            print("   次は、プロジェクト直下で以下を実行してください:")
            print("     export PYTHONPATH=\"$PWD/src\"")
        print("     python -m autotest validate")
        print("")
        print("   操作画面を開く場合は run_gui.bat を実行してください。")
    else:
        print(" 結論: この環境では実行できません。上の [NG] を解消してください。")
        print("")
        print(" 不足パッケージの導入方法:")
        print("     Anaconda      : conda install openpyxl pillow pyyaml pyodbc")
        print("     通常の Python : pip install -r requirements.txt")
        print("")
        print(" Python 自体が古い場合（3.5 以前 / Python 2）は新しい環境が必要です:")
        print("     conda create -n autotest python=3.11 openpyxl pillow pyyaml pyodbc")
    print("=" * 74)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
