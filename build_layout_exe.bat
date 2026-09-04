@echo off
rem ============================================================================
rem  Layout TXT / TIF / TAR Generator - Windows EXE build
rem
rem  Usage:
rem    build_layout_exe.bat --install   Install the matching PyInstaller, then build
rem    build_layout_exe.bat             Build with an already prepared env
rem
rem  Optional:
rem    set LAYOUT_BUILD_PYTHON=C:\ProgramData\Anaconda3\python.exe
rem ============================================================================
setlocal
cd /d "%~dp0"
chcp 65001 > nul

set "BUILD_PYTHON="
if defined LAYOUT_BUILD_PYTHON set "BUILD_PYTHON=%LAYOUT_BUILD_PYTHON%"
if not defined BUILD_PYTHON if exist ".venv\Scripts\python.exe" set "BUILD_PYTHON=%CD%\.venv\Scripts\python.exe"
if not defined BUILD_PYTHON set "BUILD_PYTHON=python"

echo [1/5] Python environment
"%BUILD_PYTHON%" --version
if errorlevel 1 (
    echo.
    echo Python was not found. Run this BAT from Anaconda Prompt, or set:
    echo   set LAYOUT_BUILD_PYTHON=C:\ProgramData\Anaconda3\python.exe
    goto :error
)

"%BUILD_PYTHON%" -c "import sys; sys.exit(0 if (3, 6) <= sys.version_info[:2] <= (3, 14) else 1)"
if errorlevel 1 (
    echo.
    echo This build supports Python 3.6 through 3.14.
    echo Python 3.6-3.10 uses PyInstaller 4.10; Python 3.11-3.14 uses 6.16.0.
    goto :error
)

set "BUILD_REQUIREMENTS=requirements-build-py36.txt"
set "BUILD_PYINSTALLER_VERSION=4.10"
"%BUILD_PYTHON%" -c "import sys; sys.exit(0 if sys.version_info[:2] <= (3, 10) else 1)"
if errorlevel 1 (
    set "BUILD_REQUIREMENTS=requirements-build-modern.txt"
    set "BUILD_PYINSTALLER_VERSION=6.16.0"
)
echo Build profile: PyInstaller %BUILD_PYINSTALLER_VERSION% ^(%BUILD_REQUIREMENTS%^)

if /I "%~1"=="--install" (
    echo.
    echo [2/5] Installing pinned build dependency
    "%BUILD_PYTHON%" -m pip install -r "%BUILD_REQUIREMENTS%"
    if errorlevel 1 goto :error
) else (
    echo.
    echo [2/5] Checking pinned build dependency
    "%BUILD_PYTHON%" -c "import PyInstaller, sys; print('PyInstaller ' + PyInstaller.__version__); sys.exit(0 if PyInstaller.__version__ == sys.argv[1] else 1)" "%BUILD_PYINSTALLER_VERSION%"
    if errorlevel 1 (
        echo.
        echo PyInstaller %BUILD_PYINSTALLER_VERSION% is not installed in this Python environment.
        echo Run: build_layout_exe.bat --install
        goto :error
    )
)

rem Python 3.5+ already includes typing in the standard library. Some old
rem Anaconda environments also contain the obsolete backport distribution,
rem which PyInstaller intentionally refuses to build with. Remove only that
rem distribution from the explicitly selected build interpreter.
"%BUILD_PYTHON%" -m pip show typing >nul 2>nul
if not errorlevel 1 (
    echo.
    echo Removing obsolete typing backport that conflicts with PyInstaller...
    "%BUILD_PYTHON%" -m pip uninstall -y typing
    if errorlevel 1 (
        echo.
        echo Could not remove the obsolete typing backport automatically.
        echo Do not use conda remove; old conda solvers may fail before removal.
        echo Run this with the selected Python, then retry:
        echo   "%BUILD_PYTHON%" -m pip uninstall -y typing
        goto :error
    )
    "%BUILD_PYTHON%" -m pip show typing >nul 2>nul
    if not errorlevel 1 (
        echo.
        echo The obsolete typing backport is still installed.
        echo Do not use conda remove. Check this command's full output:
        echo   "%BUILD_PYTHON%" -m pip uninstall -y typing
        goto :error
    )
)

echo.
echo [3/5] Checking runtime dependencies
set "PYTHONPATH=%CD%\src"
"%BUILD_PYTHON%" layout_txt_exe.py --smoke-test
if errorlevel 1 (
    echo.
    echo openpyxl, Pillow, tkinter, or the Layout generator could not be imported.
    echo Run setup_windows.bat, then retry this build.
    goto :error
)

echo.
echo [4/5] Building onedir Windows application
"%BUILD_PYTHON%" -m PyInstaller --noconfirm --clean layout_txt.spec
if errorlevel 1 goto :error

echo.
echo [5/5] Verifying packaged application
if not exist "dist\LayoutTxtGenerator\LayoutTxtGenerator.exe" (
    echo EXE was not created at the expected path.
    goto :error
)
start "" /wait "dist\LayoutTxtGenerator\LayoutTxtGenerator.exe" --smoke-test
if errorlevel 1 (
    echo Packaged application smoke test failed.
    goto :error
)

echo.
echo ============================================================================
echo Build completed successfully.
echo Distribute this entire folder, not only the EXE:
echo   %CD%\dist\LayoutTxtGenerator
echo Start application:
echo   %CD%\dist\LayoutTxtGenerator\LayoutTxtGenerator.exe
echo ============================================================================
exit /b 0

:error
echo.
echo Build failed. Check the message above.
exit /b 1
