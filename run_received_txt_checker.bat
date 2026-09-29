@echo off
rem 受領TXT確認専用アプリを起動する。
setlocal
cd /d "%~dp0"
chcp 65001 > nul
set "CHECKER_PYTHON=python"
if exist ".venv\Scripts\python.exe" set "CHECKER_PYTHON=%CD%\.venv\Scripts\python.exe"
"%CHECKER_PYTHON%" received_txt_checker.py %*
set "CHECKER_EXIT=%ERRORLEVEL%"
if not "%CHECKER_EXIT%"=="0" (
    echo.
    echo 起動できませんでした。Python・tkinter・openpyxl を確認してください。
    pause
)
exit /b %CHECKER_EXIT%
