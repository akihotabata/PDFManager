@echo off
setlocal
cd /d "%~dp0"
cd ..
if not exist ".venv" ( py -m venv .venv )
call .venv\Scripts\activate.bat
python -m pip install -U pip >nul
pip install -r tools\requirements.txt
start "" ".venv\Scripts\pythonw.exe" "src\pdf_manager_app.py"
endlocal
