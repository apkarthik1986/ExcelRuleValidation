@echo off
REM Setup script for Excel Rule Validation System

echo ================================================
echo Excel Rule Validation System - Setup
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

echo Installing required packages...
echo.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo ================================================
echo Setup completed successfully!
echo ================================================
echo.
echo To start the GUI application, run: start_gui.bat
echo To use the CLI, run: python cli_app.py --help
echo.
pause
