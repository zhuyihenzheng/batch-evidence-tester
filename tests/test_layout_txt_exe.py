# -*- coding: utf-8 -*-

import sys

import layout_txt_exe


def test_version_does_not_import_gui(capsys):
    assert layout_txt_exe.main(["--version"]) == 0
    assert capsys.readouterr().out.strip() == "LayoutTxtGenerator 0.3.2"


def test_prepare_source_path_is_idempotent():
    source_path = str(layout_txt_exe.Path(layout_txt_exe.__file__).resolve().parent / "src")
    before = sys.path.count(source_path)

    layout_txt_exe._prepare_source_path()
    after_first = sys.path.count(source_path)
    assert after_first == max(1, before)

    layout_txt_exe._prepare_source_path()
    assert sys.path.count(source_path) == after_first
