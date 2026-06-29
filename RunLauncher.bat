@echo off
setlocal
cd /d "%~dp0"

if exist "venv\Scripts\pythonw.exe" (
    "venv\Scripts\pythonw.exe" launcher.py
    exit /b %ERRORLEVEL%
)

if exist ".venv\Scripts\pythonw.exe" (
    ".venv\Scripts\pythonw.exe" launcher.py
    exit /b %ERRORLEVEL%
)

where pyw.exe >nul 2>nul
if not errorlevel 1 (
    pyw.exe launcher.py
    exit /b 0
)

if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" launcher.py
    exit /b %ERRORLEVEL%
)

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" launcher.py
    exit /b %ERRORLEVEL%
)

python launcher.py
