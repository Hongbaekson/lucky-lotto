@echo off
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 goto failed
)
".venv\Scripts\python.exe" -c "import openpyxl, numpy, scipy" >nul 2>&1
if errorlevel 1 (
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto failed
)
".venv\Scripts\python.exe" lotto_gui.py
if errorlevel 1 goto failed
exit /b 0
:failed
echo.
echo Could not start Lucky Lotto. Check the error above.
pause
exit /b 1
