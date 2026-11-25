@echo off
cd /d "%~dp0"

if not exist ".venv" (
    echo Creating local Python environment...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

python -c "import flask, openpyxl, waitress" 2>nul
if errorlevel 1 (
    echo Installing dependencies ^(flask, openpyxl, waitress^)...
    pip install --quiet flask openpyxl waitress
)

python AppUpdate.py
pause
