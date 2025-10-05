@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0\.."
if not exist "venv" (
  echo [setup] Create venv
  py -m venv venv
)
call venv\Scripts\activate.bat
python -m pip install -U pip >nul
pip install -r tools\requirements.txt
echo [run] Launching app...
start "" "%cd%\venv\Scripts\pythonw.exe" "%cd%\src\pdf_manager_app.py"
endlocal
