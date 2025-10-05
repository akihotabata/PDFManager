@echo off
:: ------------------------------------------------------------
:: PDFManager Portable Startup Script
:: - 事前にPythonポータブル環境が同一階層に存在する前提
:: - 例: PDFManager/
::        ├─ tools/
::        ├─ src/
::        └─ python-portable/   ← ここに python.exe がある
:: ------------------------------------------------------------

setlocal
cd /d %~dp0
cd ..

echo.
echo =============================================
echo  PDFManager Portable Edition
echo =============================================

:: Pythonポータブルパスを指定
set PYTHON_PATH=%~dp0..\python-portable\python.exe

if not exist "%PYTHON_PATH%" (
    echo [エラー] python-portable\python.exe が見つかりません。
    echo 同一フォルダ階層にPythonのポータブル版を配置してください。
    pause
    exit /b
)

:: 依存パッケージをインストール
echo 依存パッケージを確認中...
"%PYTHON_PATH%" -m pip install --no-warn-script-location -r tools\requirements.txt

:: アプリ起動
echo.
echo アプリを起動します...
"%PYTHON_PATH%" src\pdf_merger_app.py

echo.
pause
endlocal
