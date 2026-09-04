# -*- coding: utf-8 -*-

import sys
import tarfile
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
    assert "Do not use conda remove" in script


def test_build_script_supports_onedir_and_onefile_modes():
    root = Path(layout_txt_exe.__file__).resolve().parent
    script = (root / "build_layout_exe.bat").read_text()
    spec = (root / "layout_txt.spec").read_text()
    assert "--onefile" in script
    assert "--onedir" in script
    assert "LAYOUT_BUILD_MODE=%BUILD_MODE%" in script
    assert 'os.environ.get("LAYOUT_BUILD_MODE", "onedir")' in spec
    assert 'if build_mode == "onefile"' in spec


def test_smoke_test_exercises_excel_txt_tif_and_tar(monkeypatch, tmp_path):
    created_temp = tmp_path / "smoke"
    created_temp.mkdir()

    monkeypatch.setattr(
        layout_txt_exe.tempfile, "mkdtemp",
        lambda prefix: str(created_temp))
    monkeypatch.setattr(
        layout_txt_exe, "_cleanup_smoke_directory",
        lambda path: True)

    assert layout_txt_exe._functional_smoke_test() == 0
    assert (created_temp / "output" / "1001.txt").is_file()
    assert (created_temp / "output" / "1001.tif").is_file()
    package = created_temp / "output" / "smoke_package.tar"
    assert package.is_file()
    with tarfile.open(str(package), "r") as archive:
        assert archive.getnames() == ["1001.txt", "1001.tif"]


def test_smoke_cleanup_retries_transient_windows_file_lock(monkeypatch, tmp_path):
    created_temp = tmp_path / "locked"
    created_temp.mkdir()
    attempts = []
    real_rmtree = layout_txt_exe.shutil.rmtree

    def flaky_rmtree(path):
        attempts.append(path)
        if len(attempts) == 1:
            raise PermissionError(32, "file is being used", "smoke.xlsx")
        real_rmtree(path)

    monkeypatch.setattr(layout_txt_exe.shutil, "rmtree", flaky_rmtree)
    monkeypatch.setattr(layout_txt_exe.time, "sleep", lambda delay: None)

    assert layout_txt_exe._cleanup_smoke_directory(
        created_temp, attempts=2, delay=0) is True
    assert len(attempts) == 2
    assert not created_temp.exists()


def test_smoke_cleanup_does_not_fail_build_for_persistent_lock(
        monkeypatch, tmp_path):
    created_temp = tmp_path / "locked"
    created_temp.mkdir()

    def locked_rmtree(path):
        raise PermissionError(32, "file is being used", "smoke.xlsx")

    monkeypatch.setattr(layout_txt_exe.shutil, "rmtree", locked_rmtree)
    monkeypatch.setattr(layout_txt_exe.time, "sleep", lambda delay: None)

    assert layout_txt_exe._cleanup_smoke_directory(
        created_temp, attempts=2, delay=0) is False


def test_spec_excludes_unrelated_broken_anaconda_gevent_packages():
    root = Path(layout_txt_exe.__file__).resolve().parent
    spec = (root / "layout_txt.spec").read_text()
    assert '"gevent"' in spec
    assert '"greenlet"' in spec


def test_spec_does_not_bundle_large_optional_anaconda_packages():
    root = Path(layout_txt_exe.__file__).resolve().parent
    spec = (root / "layout_txt.spec").read_text()
    assert "collect_submodules" not in spec
    for package in ("numpy", "pandas", "scipy", "matplotlib", "PyQt5"):
        assert '"%s"' % package in spec
