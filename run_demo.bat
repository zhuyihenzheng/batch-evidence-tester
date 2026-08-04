@echo off
rem ============================================================================
rem  サンドボックス実行（実環境に一切触れない動作確認）
rem    被テスト .exe    → demo\fake_batch.py
rem    SQL Server       → fixtures\ の CSV
rem    入出力フォルダ   → .\sandbox\
rem  ツール導入直後に、Excel 成果物がどう出るかを確認するために使う。
rem ============================================================================
setlocal
cd /d "%~dp0"
chcp 65001 > nul

rem  .venv があればそれを使い、無ければ PATH 上の python を使う（Anaconda 想定）
if exist ".venv\Scripts\python.exe" (
    set PY=.venv\Scripts\python.exe
) else (
    where python > nul 2>&1
    if errorlevel 1 (
        echo python が見つかりません。「Anaconda Prompt」から実行するか、
        echo setup_windows.bat を先に実行してください。
        pause
        exit /b 2
    )
    set PY=python
)

if exist "sandbox" rmdir /s /q sandbox

set PYTHONPATH=%~dp0src
%PY% -m autotest run --config config\settings.demo.yaml --offline
echo.
echo output フォルダの TestEvidence_*.xlsx を開いて確認してください。
echo   TC001 / TC002 … OK 判定の見本
echo   TC003        … NG 判定と差分表の見本（意図的に期待値をずらしてあります）
pause
