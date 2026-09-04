@echo off
rem ============================================================================
rem  Layout TXT / TIF / TAR Generator - Windows EXE build
rem
rem  Usage:
rem    build_layout_exe.bat --install            Install dependencies; build onedir
rem    build_layout_exe.bat --onefile            Build one distributable EXE
rem    build_layout_exe.bat --onefile --install  Combine options in any order
rem    build_layout_exe.bat --onedir             Explicitly build the default mode
rem
rem  Optional:
rem    set LAYOUT_BUILD_PYTHON=C:\ProgramData\Anaconda3\python.exe
rem ============================================================================
setlocal
cd /d "%~dp0"
chcp 65001 > nul

set "INSTALL_DEPS=0"
set "BUILD_MODE=onedir"

:parse_args
if "%~1"=="" goto :args_done
if /I "%~1"=="--install" (
    set "INSTALL_DEPS=1"
) else if /I "%~1"=="--onefile" (
    set "BUILD_MODE=onefile"
) else if /I "%~1"=="--onedir" (
    set "BUILD_MODE=onedir"
) else (
    echo Unknown option: %~1
    echo Usage: build_layout_exe.bat [--install] [--onedir ^| --onefile]
    goto :error
)
shift
goto :parse_args

:args_done
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
echo Package mode: %BUILD_MODE%

if "%INSTALL_DEPS%"=="1" (
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
    echo Runtime dependency check or Excel/TXT/TIF/TAR smoke test failed.
    echo Review the traceback above. If an import is missing, run setup_windows.bat.
    goto :error
)

echo.
echo [4/5] Building %BUILD_MODE% Windows application
set "LAYOUT_BUILD_MODE=%BUILD_MODE%"
"%BUILD_PYTHON%" -m PyInstaller --noconfirm --clean layout_txt.spec
if errorlevel 1 goto :error

echo.
echo [5/5] Verifying packaged application
if /I "%BUILD_MODE%"=="onefile" (
    set "PACKAGED_EXE=%CD%\dist\LayoutTxtGenerator.exe"
) else (
    set "PACKAGED_EXE=%CD%\dist\LayoutTxtGenerator\LayoutTxtGenerator.exe"
)
if not exist "%PACKAGED_EXE%" (
    echo EXE was not created at the expected path.
    goto :error
)
start "" /wait "%PACKAGED_EXE%" --smoke-test
if errorlevel 1 (
    echo Packaged application smoke test failed.
    goto :error
)

echo.
echo ============================================================================
echo Build completed successfully.
if /I "%BUILD_MODE%"=="onefile" (
    echo Distribute this single EXE:
) else (
    echo Distribute this entire folder, not only the EXE:
)
echo Start application:
echo   %PACKAGED_EXE%
echo ============================================================================
exit /b 0

:error
echo.
echo Build failed. Check the message above.
exit /b 1
