@echo off
rem Excel の I/J/K 列から Layout 取込用 TXT を生成する画面を開く。
setlocal
cd /d "%~dp0"
chcp 65001 > nul

set PY=
if exist ".venv\Scripts\pythonw.exe" set PY=.venv\Scripts\pythonw.exe
if not defined PY if exist ".venv\Scripts\python.exe" set PY=.venv\Scripts\python.exe
if not defined PY (
    where pythonw > nul 2>&1
    if not errorlevel 1 set PY=pythonw
)
if not defined PY (
    where python > nul 2>&1
    if not errorlevel 1 set PY=python
)

if not defined PY (
    echo Python が見つかりません。check_env.py を実行して環境を確認してください。
    pause
    exit /b 2
)

set PYTHONPATH=%~dp0src
%PY% -m autotest.layout_txt --gui
if errorlevel 1 (
    echo.
    echo Layout TXT 生成画面を開けませんでした。
    pause
    exit /b 1
)
exit /b 0
