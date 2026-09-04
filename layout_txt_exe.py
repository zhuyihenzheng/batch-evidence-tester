# -*- coding: utf-8 -*-
"""Windows EXE entry point for the Layout TXT/TIF/TAR generator."""

import sys
import tarfile
import tempfile
from pathlib import Path


APP_VERSION = "0.4.0"


def _prepare_source_path():
    """Allow the entry point to run directly before it is frozen."""
    if not getattr(sys, "frozen", False):
        source_dir = Path(__file__).resolve().parent / "src"
        source_text = str(source_dir)
        if source_text not in sys.path:
            sys.path.insert(0, source_text)


def _functional_smoke_test():
    """Exercise the packaged Excel -> TXT/TIF/TAR generation path."""
    from openpyxl import Workbook
    from PIL import Image

    from autotest.layout_txt import generate_layout_txt

    with tempfile.TemporaryDirectory(prefix="layout_txt_smoke_") as temp_dir:
        root = Path(temp_dir)
        workbook_path = root / "smoke.xlsx"
        output_dir = root / "output"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append([
            "FORM_ID", "LAYOUT_ID", "FORM_NAME", "RESERVED_1",
            "RESERVED_2", "RESERVED_3", "RESERVED_4", "ITEM_NAME",
            "ELEMENT_DATA_TYPE_NAME", "ELEMENT_IME_NAME",
            "MAX_NUM_DIGITS", "ELEMENT_ID",
        ])
        sheet.append([
            1001, 1, "Smoke", "", "", "", "", "Item",
            "文字列", "半角英数", 10, 9001,
        ])
        workbook.save(str(workbook_path))
        if hasattr(workbook, "close"):
            workbook.close()

        result = generate_layout_txt(
            workbook_path, output_dir, error_patterns="none",
            generate_tif=True, create_tar=True, tar_name="smoke_package")
        if len(result.txt_files) != 1 or len(result.tif_files) != 1:
            raise RuntimeError("Smoke test did not generate one TXT/TIF pair")
        with Image.open(str(result.tif_files[0])) as image:
            image.load()
        with tarfile.open(str(result.tar_file), "r") as archive:
            if sorted(archive.getnames()) != sorted(result.archive_members):
                raise RuntimeError("Smoke TAR members do not match generated files")

    return 0


def _smoke_test():
    """Exercise GUI imports and actual generation without opening a window."""
    import tkinter  # noqa: F401

    from autotest import layout_txt_gui  # noqa: F401

    return _functional_smoke_test()


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
