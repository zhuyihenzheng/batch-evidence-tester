# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


project_root = Path(SPECPATH)
sys.path.insert(0, str(project_root))

# Anaconda 5.2 backported a private sysconfig function with an incompatible
# required argument.  Patch it before PyInstaller 4.10 loads its built-in
# distutils and sysconfig hooks.
from pyinstaller_compat import patch_anaconda_sysconfig

patch_anaconda_sysconfig()

# PyInstaller's built-in openpyxl/Pillow hooks already collect their required
# modules. Collecting every submodule here also follows optional imports into
# large packages present in Anaconda (NumPy/Pandas/MKL/Qt), inflating this
# small application to hundreds of megabytes.
hidden_imports = []
data_files = collect_data_files("openpyxl")

block_cipher = None
build_mode = os.environ.get("LAYOUT_BUILD_MODE", "onedir").strip().lower()
if build_mode not in ("onedir", "onefile"):
    raise ValueError("LAYOUT_BUILD_MODE must be onedir or onefile: %s" % build_mode)


a = Analysis(
    [str(project_root / "layout_txt_exe.py")],
    pathex=[str(project_root / "src")],
    binaries=[],
    datas=data_files,
    hiddenimports=hidden_imports,
    hookspath=[],
    runtime_hooks=[],
    # The standalone Layout generator does not use these packages. In old
    # Anaconda installations gevent/greenlet metadata is often incomplete;
    # allowing PyInstaller to discover it loads hook-gevent and aborts the
    # otherwise unrelated build.
    excludes=[
        "mss", "pyodbc", "yaml", "gevent", "greenlet",
        "numpy", "pandas", "scipy", "matplotlib",
        "IPython", "jupyter", "notebook",
        "PyQt5", "PyQt6", "PySide2", "PySide6",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

if build_mode == "onefile":
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name="LayoutTxtGenerator",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
        version=str(project_root / "layout_txt_version_info.txt"),
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="LayoutTxtGenerator",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
        version=str(project_root / "layout_txt_version_info.txt"),
    )

    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name="LayoutTxtGenerator",
    )
