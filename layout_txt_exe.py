# -*- coding: utf-8 -*-
"""Windows EXE entry point for the Layout TXT/TIF/TAR generator."""

import sys
from pathlib import Path


APP_VERSION = "0.3.5"


def _prepare_source_path():
    """Allow the entry point to run directly before it is frozen."""
    if not getattr(sys, "frozen", False):
        source_dir = Path(__file__).resolve().parent / "src"
        source_text = str(source_dir)
        if source_text not in sys.path:
            sys.path.insert(0, source_text)


def _smoke_test():
    """Import every runtime dependency without opening a GUI window."""
    import tkinter  # noqa: F401

    import openpyxl  # noqa: F401
    from PIL import Image  # noqa: F401

    from autotest import layout_txt  # noqa: F401
    from autotest import layout_txt_gui  # noqa: F401

    return 0


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    _prepare_source_path()

    if "--version" in args:
        print("LayoutTxtGenerator %s" % APP_VERSION)
        return 0
    if "--smoke-test" in args:
        return _smoke_test()

    from autotest.layout_txt_gui import main as gui_main

    return gui_main()


if __name__ == "__main__":
    sys.exit(main())
