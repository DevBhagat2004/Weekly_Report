#!/bin/bash

# Automated Weekly Reports - Shell Script
# This script runs the weekly report generator

echo "========================================"
echo "  Automated Weekly Reports Generator"
echo "========================================"
echo

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "ERROR: Python is not installed"
        echo "Please install Python 3.7 or higher"
        exit 1
    else
        PYTHON_CMD="python"
    fi
else
    PYTHON_CMD="python3"
fi

echo "Using Python: $PYTHON_CMD"

# Check Python version
PYTHON_VERSION=$($PYTHON_CMD --version 2>&1 | cut -d' ' -f2 | cut -d'.' -f1,2)
echo "Python version: $PYTHON_VERSION"

# Check if required packages are installed
echo "Checking dependencies..."
$PYTHON_CMD -c "import pandas, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing required packages..."
    pip3 install -r requirements.txt 2>/dev/null || pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to install dependencies"
        echo "Try running: pip3 install pandas openpyxl"
        exit 1
    fi
fi

# Run the report generator
echo
echo "Running weekly report generator..."
echo
$PYTHON_CMD weekly_report_generator.py

if [ $? -eq 0 ]; then
    echo
    echo "SUCCESS: Weekly report generated successfully! ✅"
    echo "Check the output Excel file in the current directory"
    
    # List generated files
    echo
    echo "Generated files:"
    ls -la *.xlsx 2>/dev/null | head -5
    ls -la *.log 2>/dev/null | head -3
else
    echo
    echo "ERROR: Report generation failed ❌"
    echo "Check the log file for details: weekly_report.log"
    exit 1
fi

echo
