@echo off
REM Automated Weekly Reports - Windows Batch Script
REM This script runs the weekly report generator

echo ========================================
echo   Automated Weekly Reports Generator
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher
    pause
    exit /b 1
)

REM Check if required packages are installed
echo Checking dependencies...
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

REM Run the report generator
echo.
echo Running weekly report generator...
echo.
python weekly_report_generator.py

if errorlevel 1 (
    echo.
    echo ERROR: Report generation failed
    echo Check the log file for details: weekly_report.log
) else (
    echo.
    echo SUCCESS: Weekly report generated successfully!
    echo Check the output Excel file in the current directory
)

echo.
pause
