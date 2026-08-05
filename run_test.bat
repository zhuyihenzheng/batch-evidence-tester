@echo off
rem ============================================================================
rem  自動テスト実行
rem    run_test.bat                    … 全ケース
rem    run_test.bat --case TC001_normal … ケース指定
rem    run_test.bat --tag 正常系         … タグ指定
rem    run_test.bat --dry-run           … .exe も DB も触らず流れだけ確認
rem
rem  設定ファイルはツール側が決める（config\settings.local.yaml があれば
rem  そちらを優先し、無ければ settings.yaml）。ここで --config を固定すると
rem  local 設定が読まれず、テンプレートの exe や DB を見に行ってしまう。
rem  別の設定で動かしたい場合は  run_test.bat --config <パス>  と渡す。
rem
rem  終了コード:
rem    0 = 全 OK / 1 = NG あり / 2 = 設定エラー / 3 = 人の確認待ち
rem  （タスクスケジューラや Jenkins からはこの終了コードで判定できる）
rem ============================================================================
setlocal
cd /d "%~dp0"
chcp 65001 > nul

rem  .venv があればそれを使い、無ければ PATH 上の python を使う。
rem  Anaconda 環境では .venv を作らずそのまま実行できるようにするため。
if exist ".venv\Scripts\python.exe" (
    set PY=.venv\Scripts\python.exe
) else (
    where python > nul 2>&1
    if errorlevel 1 (
        echo python が見つかりません。
        echo   Anaconda をお使いの場合は「Anaconda Prompt」から実行してください。
        echo   通常の Python の場合は setup_windows.bat を先に実行してください。
        pause
        exit /b 2
    )
    set PY=python
)

set PYTHONPATH=%~dp0src
%PY% -m autotest run %*
set RESULT=%errorlevel%

echo.
if "%RESULT%"=="0" (
    echo 総合判定: OK
) else if "%RESULT%"=="1" (
    echo 総合判定: NG  ―  出力された Excel の「サマリ」シートを確認してください。
) else if "%RESULT%"=="3" (
    echo 総合判定: 要確認  ―  Excel の黄色い欄に確認結果を記入してください。
) else if "%RESULT%"=="2" (
    echo 設定エラーで実行できませんでした。
) else (
    echo 想定外の終了コード: %RESULT%
)

rem 対話実行時のみ一時停止する（スケジューラ起動時は止めない）
echo %cmdcmdline% | find /i "%~0" > nul
if not errorlevel 1 pause
exit /b %RESULT%
