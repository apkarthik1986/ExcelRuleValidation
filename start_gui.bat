@echo off
REM Excel Rule Validation System - GUI Launcher
REM This batch file launches the GUI application

echo ================================================
echo Excel Rule Validation System
echo ================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://www.python.org/
    pause
    exit /b 1
)

echo Starting GUI Application...
echo.

python gui_app.py

if errorlevel 1 (
    echo.
    echo ERROR: Application failed to start
    echo Make sure all dependencies are installed: pip install -r requirements.txt
    pause
)
