@echo off
REM Excel Rule Validation System - GUI Launcher
REM This batch file launches the GUI application using the project venv

echo ================================================
echo Excel Rule Validation System
echo ================================================
echo.

REM Prefer .venv311 (Python 3.11) then .venv, else system Python
if exist ".venv311\Scripts\activate.bat" (
    echo Activating .venv311
    call ".venv311\Scripts\activate.bat"
) else if exist ".venv\Scripts\activate.bat" (
    echo Activating .venv
    call ".venv\Scripts\activate.bat"
) else (
    echo No virtual environment found; falling back to system Python
)

REM Ensure logs directory exists
if not exist "logs" mkdir "logs"

set LOGFILE=logs\gui.log
echo Starting GUI Application (logging to %LOGFILE%)...

python gui_app.py > "%LOGFILE%" 2>&1

if errorlevel 1 (
    echo.
    echo ERROR: Application failed to start. See %LOGFILE% for details.
    echo Make sure dependencies are installed: python -m pip install -r requirements.txt
    pause
)

echo Application exited. Logs are available at %LOGFILE%
