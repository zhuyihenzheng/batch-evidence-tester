# -*- coding: utf-8 -*-
"""元の生成GUI・DBツールを含まない、単独実行用ソース配布を作る。"""

import argparse
import zipfile
from pathlib import Path


def package(output):
    root = Path(__file__).resolve().parent.parent
    files = [
        "received_txt_checker.py", "received_txt_checker.spec",
        "run_received_txt_checker.bat", "build_received_txt_checker.bat",
        "requirements-received-txt-checker.txt", "requirements-build-py36.txt",
        "requirements-build-modern.txt", "pyinstaller_compat.py", "LICENSE",
        "src/autotest/layout_inspect.py", "src/autotest/layout_inspect_gui.py",
        "src/autotest/layout_txt.py", "src/autotest/layout_naming.py",
    ]
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(str(output), "w", zipfile.ZIP_DEFLATED) as archive:
        for name in files:
            archive.write(str(root / name), "ReceivedTxtChecker/" + name)
        archive.writestr("ReceivedTxtChecker/src/autotest/__init__.py", "")
        archive.write(str(root / "received_txt_checker_README.md"), "ReceivedTxtChecker/README.md")
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="受領TXT確認アプリの単独ソース配布を作成")
    parser.add_argument("--out", default="output/ReceivedTxtChecker-source.zip")
    print(package(parser.parse_args().out))
