@echo off
setlocal
cd /d "%~dp0"
cd ..
if not exist ".venv" ( py -m venv .venv )
call .venv\Scripts\activate.bat
pip install -r tools\requirements.txt
pip install pyinstaller
pyinstaller --noconsole --onefile --name "PDF整理ツール" --icon=docs\icon.ico ^
  --hidden-import fitz --hidden-import pymupdf src\pdf_manager_app.py
echo 出力: dist\PDF整理ツール.exe
pause
