@echo off
set "USER_APP_DIR=%~dp0"
set "BACKEND_DIR=C:\Users\Nitin Maheshwari\Desktop\Motifzone Hiring Automation Tool"
set "LOG_FILE=%USER_APP_DIR%Hiring Automation Tool Launch Log.txt"

if not exist "%BACKEND_DIR%\.venv\Scripts\pythonw.exe" (
    echo Hiring Automation Tool backend was not found.
    echo Please contact Motifzone support.
    echo Backend missing: %BACKEND_DIR% > "%LOG_FILE%"
    pause
    exit /b 1
)

set "CV_AGENT_OUTPUT=%USER_APP_DIR%Candidate Details.xlsx"
set "CV_AGENT_ATTACHMENTS_DIR=%USER_APP_DIR%Downloaded CVs"

if not exist "%CV_AGENT_ATTACHMENTS_DIR%" mkdir "%CV_AGENT_ATTACHMENTS_DIR%"

start "Hiring Automation Tool" "%BACKEND_DIR%\.venv\Scripts\pythonw.exe" "%BACKEND_DIR%\cv_agent_ui.py"

if errorlevel 1 (
    echo Unable to launch Hiring Automation Tool. > "%LOG_FILE%"
    echo Backend: %BACKEND_DIR% >> "%LOG_FILE%"
    echo Please contact Motifzone support. >> "%LOG_FILE%"
    start notepad "%LOG_FILE%"
)
