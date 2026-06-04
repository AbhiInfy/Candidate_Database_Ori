@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Python virtual environment not found.
    echo Please run setup first in VS Code: python -m venv .venv
    pause
    exit /b 1
)

echo Starting Hiring Automation Tool...
echo.
".venv\Scripts\python.exe" ".\imap_cv_agent.py" --imap-server "mail.emotifzone.com" --username "careers@emotifzone.com" --folder "INBOX"

echo.
echo Finished. Press any key to close this window.
pause >nul
