@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0\.."
if not exist "venv" py -m venv venv
call venv\Scripts\activate.bat
python -m pip install -U pip >nul
pip install -r tools\requirements.txt
pip install pyinstaller
set ICON=docs\icon.ico
if not exist %ICON% set ICON=
echo [build] Building PDFManager.exe
pyinstaller --noconsole --onefile --name "PDFManager" --hidden-import fitz --hidden-import pymupdf %ICON% src\pdf_manager_app.py
echo [done] dist\PDFManager.exe
pause
endlocal
