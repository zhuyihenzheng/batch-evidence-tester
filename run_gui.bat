@echo off
rem ============================================================================
rem  操作画面を開く
rem    ケース一覧を見ながら、全件 / タグ / 個別に実行できる。
rem    ケースの新規作成・複製、手動実施ケースの証跡採取もここから行える。
rem
rem  無人実行（タスクスケジューラ等）は従来どおり run_test.bat を使うこと。
rem  この画面は対話操作用。
rem ============================================================================
setlocal
cd /d "%~dp0"
chcp 65001 > nul

rem  pythonw.exe を優先する（黒いコンソール窓を出さないため）。
rem  無ければ python.exe で起動する（窓は出るが動作は同じ）。
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
    echo python が見つかりません。
    echo   Anaconda をお使いの場合は「Anaconda Prompt」から実行してください。
    echo   通常の Python の場合は setup_windows.bat を先に実行してください。
    pause
    exit /b 2
)

set PYTHONPATH=%~dp0src
%PY% -m autotest gui %*

rem  pythonw は即座に戻るので、起動に失敗した場合だけここに意味がある。
rem  画面が出ないときは python.exe に切り替えてエラーを確認すること:
rem      set PYTHONPATH=src
rem      .venv\Scripts\python.exe -m autotest gui
if errorlevel 1 (
    echo.
    echo 画面を開けませんでした。次のコマンドで詳しいエラーを確認できます:
    echo   set PYTHONPATH=src
    echo   python -m autotest gui
    pause
    exit /b 1
)
exit /b 0
