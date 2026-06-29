@echo off
setlocal
cd /d "%~dp0"

if exist "venv\Scripts\pythonw.exe" (
    "venv\Scripts\pythonw.exe" launcher.py
    exit /b %ERRORLEVEL%
)

if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" launcher.py
    exit /b %ERRORLEVEL%
)

python launcher.py
