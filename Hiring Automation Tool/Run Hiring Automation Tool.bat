@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
    echo Python virtual environment not found.
    echo Please run setup first in VS Code: python -m venv .venv
    pause
    exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" ".\cv_agent_ui.py"
