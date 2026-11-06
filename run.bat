@echo off
rem Marksmith - one-click start (Windows).
rem First run: creates .venv, installs the Python packages and downloads Chromium (~150 MB).
cd /d "%~dp0"
where python >nul 2>nul || (echo Python 3.10+ not found. Install it from python.org and tick "Add python to PATH". & pause & exit /b 1)
if not exist ".venv\Scripts\python.exe" python -m venv .venv || (pause & exit /b 1)
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt || (pause & exit /b 1)
".venv\Scripts\python.exe" -m playwright install chromium || (pause & exit /b 1)
".venv\Scripts\python.exe" app.py
pause
