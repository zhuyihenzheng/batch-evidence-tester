# -*- coding: utf-8 -*-

import sys
from pathlib import Path

import layout_txt_exe


def test_version_does_not_import_gui(capsys):
    assert layout_txt_exe.main(["--version"]) == 0
    assert capsys.readouterr().out.strip() == "LayoutTxtGenerator 0.4.0"


def test_prepare_source_path_is_idempotent():
    source_path = str(layout_txt_exe.Path(layout_txt_exe.__file__).resolve().parent / "src")
    before = sys.path.count(source_path)

    layout_txt_exe._prepare_source_path()
    after_first = sys.path.count(source_path)
    assert after_first == max(1, before)

    layout_txt_exe._prepare_source_path()
    assert sys.path.count(source_path) == after_first


def test_legacy_and_modern_pyinstaller_versions_are_pinned():
    root = Path(layout_txt_exe.__file__).resolve().parent
    assert (root / "requirements-build-py36.txt").read_text().splitlines()[-1] == (
        "pyinstaller==4.10")
    assert (root / "requirements-build-modern.txt").read_text().splitlines()[-1] == (
        "pyinstaller==6.16.0")


def test_build_script_selects_modern_profile_for_new_python():
    root = Path(layout_txt_exe.__file__).resolve().parent
    script = (root / "build_layout_exe.bat").read_text()
    assert "(3, 14)" in script
    assert "requirements-build-modern.txt" in script
    assert "BUILD_PYINSTALLER_VERSION=6.16.0" in script


def test_build_script_removes_obsolete_typing_backport():
    root = Path(layout_txt_exe.__file__).resolve().parent
    script = (root / "build_layout_exe.bat").read_text()
    assert "pip show typing" in script
    assert "pip uninstall -y typing" in script
    assert "conda remove typing" in script
