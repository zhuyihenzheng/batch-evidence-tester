@echo off
rem ============================================================================
rem  AUTO_TEST_BATCH  Windows セットアップ
rem    このファイルをダブルクリック、またはコマンドプロンプトから実行する。
rem    前提: Python 3.9 以降がインストール済み（py コマンドが使えること）
rem ============================================================================
setlocal
cd /d "%~dp0"
chcp 65001 > nul

echo [1/4] Python の確認
py -3 --version
if errorlevel 1 (
    echo.
    echo   Python が見つかりません。https://www.python.org/downloads/windows/ から
    echo   3.9 以降をインストールし、"Add python.exe to PATH" にチェックを入れてください。
    goto :error
)

echo.
echo [2/4] 仮想環境の作成 .venv
if not exist ".venv" (
    py -3 -m venv .venv
    if errorlevel 1 goto :error
) else (
    echo       既存の .venv を使用します
)

echo.
echo [3/4] 依存ライブラリのインストール
rem  pip install -e . ではなく requirements.txt を使う。
rem  前者はビルド分離のため setuptools を都度取得しに行き、
rem  閉じた社内ネットワークでは失敗することがあるため。
.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
.venv\Scripts\python.exe -m pip install -r requirements.txt --quiet
if errorlevel 1 goto :error

echo.
echo [4/4] 設定の検証
set PYTHONPATH=%~dp0src
.venv\Scripts\python.exe -m autotest validate
echo.
echo ============================================================================
echo  セットアップ完了。
echo.
echo  次の手順:
echo    1. config\settings.yaml を編集
echo         batch.exe_path  … 被テスト .exe のパス
echo         paths.*         … 入出力フォルダのパス
echo         database.*      … SQL Server 接続先
echo    2. DB パスワードを環境変数に設定（設定後コンソールを開き直す）
echo         setx AUTOTEST_DB_PASSWORD "パスワード"
echo    3. Microsoft ODBC Driver 18 for SQL Server をインストール
echo         https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server
echo    4. run_test.bat を実行
echo.
echo  ※ 実環境に触れずに動作確認したい場合:  run_demo.bat
echo ============================================================================
pause
exit /b 0

:error
echo.
echo  セットアップに失敗しました。上のエラーメッセージを確認してください。
pause
exit /b 1
