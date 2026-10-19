@echo off
title Outlook Bulk Document Printer
cd /d "%~dp0"
python app.py
if %errorlevel% neq 0 (
    echo.
    echo Application exited with error code %errorlevel%.
    pause
)
