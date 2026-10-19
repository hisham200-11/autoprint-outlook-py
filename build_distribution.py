import os
import sys
import shutil
import subprocess

def build():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)
    print(f"Building standalone package in: {base_dir}")

    # Build directory
    dist_dir = os.path.join(base_dir, "dist")
    build_dir = os.path.join(base_dir, "build")

    # PyInstaller command:
    # --noconsole: clean GUI window without black cmd box
    # --name: AutoPrint-Outlook
    # --hidden-import: win32com, win32print, pypdf
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name", "AutoPrint-Outlook",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "win32print",
        "--hidden-import", "win32api",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pypdf",
        "--hidden-import", "updater",
        "app.py"
    ]

    print("Running PyInstaller...")
    res = subprocess.run(cmd, check=True)
    
    app_output_folder = os.path.join(dist_dir, "AutoPrint-Outlook")
    print(f"Build complete. Output folder: {app_output_folder}")

    # Create Uninstaller Script inside the distribution folder
    uninstaller_path = os.path.join(app_output_folder, "Uninstall_and_Delete.bat")
    uninstaller_content = """@echo off
title Uninstall AutoPrint Outlook
cls
echo ==============================================================
echo           AutoPrint Outlook - Uninstaller & Cleanup
echo ==============================================================
echo.
echo This will completely remove AutoPrint Outlook and its temporary
echo files from this computer.
echo.
set /p confirm="Are you sure you want to delete this program? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo.
    echo Uninstallation cancelled.
    pause
    exit /b
)

echo.
echo 1. Terminating any running AutoPrint processes...
taskkill /f /im AutoPrint-Outlook.exe >nul 2>&1

echo 2. Removing application files and folder...
cd /d "%~dp0.."
timeout /t 2 /nobreak >nul

rmdir /s /q "%~dp0" >nul 2>&1
if exist "%~dp0" (
    powershell -Command "Start-Sleep -Seconds 2; Remove-Item -Recurse -Force -LiteralPath '%~dp0' -ErrorAction SilentlyContinue"
)

echo.
echo ==============================================================
echo  AutoPrint Outlook has been completely deleted.
echo ==============================================================
pause
"""
    with open(uninstaller_path, "w", encoding="utf-8") as f:
        f.write(uninstaller_content)
    print(f"Created uninstaller: {uninstaller_path}")

    # Create User Guide README
    readme_path = os.path.join(app_output_folder, "HOW_TO_USE.txt")
    readme_content = """==============================================================
               AutoPrint Outlook - Quick Start Guide
==============================================================

HOW TO RUN:
1. Make sure classic Microsoft Outlook is open on your PC.
2. Double-click "AutoPrint-Outlook.exe".
3. That's it! No Python, no command-line, no setup required.

FEATURES:
- Automatically detects your Outlook Inbox and folders.
- Extracts PDFs, Images (PNG/JPG), Word docs, and Excel sheets.
- Lets you select what to print with simple checkmarks.
- Bulk prints sequentially (1-by-1) without jamming your printer.
- Automatically archives printed files into the "Printed_Archive" folder.

HOW TO UNINSTALL / DELETE:
- Double-click "Uninstall_and_Delete.bat" in this folder.
- It will completely remove the program and clean up your PC.
"""
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(readme_content)
    print(f"Created User Guide: {readme_path}")

    # Zip the distribution for 1-click sharing
    zip_output_path = os.path.join(dist_dir, "AutoPrint-Outlook-Standalone")
    print(f"Creating portable zip archive: {zip_output_path}.zip ...")
    shutil.make_archive(zip_output_path, "zip", dist_dir, "AutoPrint-Outlook")
    print(f"SUCCESS! Ready-to-ship zip created: {zip_output_path}.zip")

if __name__ == "__main__":
    build()
