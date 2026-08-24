@echo off
rem ============================================================================
rem  Layout TXT / TIF / TAR Generator - Windows EXE build
rem
rem  Usage:
rem    build_layout_exe.bat --install   Install PyInstaller 4.10, then build
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

"%BUILD_PYTHON%" -c "import sys; sys.exit(0 if (3, 6) <= sys.version_info[:2] <= (3, 10) else 1)"
if errorlevel 1 (
    echo.
    echo PyInstaller 4.10 requires Python 3.6 through 3.10 for this build.
    echo The target environment is Anaconda 5.2 / Python 3.6.5.
    goto :error
)

if /I "%~1"=="--install" (
    echo.
    echo [2/5] Installing pinned build dependency
    "%BUILD_PYTHON%" -m pip install -r requirements-build-py36.txt
    if errorlevel 1 goto :error
) else (
    echo.
    echo [2/5] Checking pinned build dependency
    "%BUILD_PYTHON%" -c "import PyInstaller, sys; print('PyInstaller ' + PyInstaller.__version__); sys.exit(0 if PyInstaller.__version__ == '4.10' else 1)"
    if errorlevel 1 (
        echo.
        echo PyInstaller 4.10 is not installed in this Python environment.
        echo Run: build_layout_exe.bat --install
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
