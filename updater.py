import os
import sys
import json
import re
import urllib.request
import zipfile
import subprocess
import time
from typing import Dict, Any, Optional

CURRENT_VERSION = "1.0.0"
GITHUB_REPO = "hisham200-11/autoprint-outlook-py"

def parse_version_tuple(v_str: str):
    # Remove leading 'v' or non-digits, split by '.'
    clean = re.sub(r'^[^\d]*', '', v_str.strip())
    parts = []
    for piece in clean.split('.'):
        digits = re.match(r'\d+', piece)
        parts.append(int(digits.group()) if digits else 0)
    return tuple(parts)

def is_version_newer(current: str, remote: str) -> bool:
    try:
        cur_tuple = parse_version_tuple(current)
        rem_tuple = parse_version_tuple(remote)
        return rem_tuple > cur_tuple
    except Exception:
        return False

def check_for_updates(repo: str = GITHUB_REPO, current_version: str = CURRENT_VERSION) -> Dict[str, Any]:
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "AutoPrint-Outlook-App"}
    )
    
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            
        latest_tag = data.get("tag_name", "")
        release_name = data.get("name", latest_tag)
        release_notes = data.get("body", "No release notes provided.")
        
        # Look for standalone zip asset
        download_url = None
        for asset in data.get("assets", []):
            name = asset.get("name", "").lower()
            if "standalone" in name or name.endswith(".zip"):
                download_url = asset.get("browser_download_url")
                break
                
        has_update = is_version_newer(current_version, latest_tag)
        
        return {
            "success": True,
            "update_available": has_update,
            "current_version": current_version,
            "latest_version": latest_tag,
            "release_name": release_name,
            "release_notes": release_notes,
            "download_url": download_url
        }
    except Exception as e:
        return {
            "success": False,
            "update_available": False,
            "current_version": current_version,
            "error": str(e)
        }

def download_and_prepare_update(download_url: str, target_dir: str, progress_callback=None) -> str:
    """
    Downloads the update zip archive into a temporary staging folder.
    Returns path to downloaded zip.
    """
    temp_dir = os.path.join(target_dir, "_update_staging")
    os.makedirs(temp_dir, exist_ok=True)
    zip_dest = os.path.join(temp_dir, "update_payload.zip")
    
    req = urllib.request.Request(
        download_url,
        headers={"User-Agent": "AutoPrint-Outlook-App"}
    )
    
    with urllib.request.urlopen(req, timeout=60) as resp, open(zip_dest, "wb") as f:
        total_len = resp.headers.get('content-length')
        total_bytes = int(total_len) if total_len else 0
        downloaded = 0
        chunk_size = 65536
        
        while True:
            chunk = resp.read(chunk_size)
            if not chunk:
                break
            f.write(chunk)
            downloaded += len(chunk)
            if progress_callback and total_bytes > 0:
                progress_callback(downloaded / total_bytes)
                
    return zip_dest

def launch_safe_updater_script(app_dir: str, update_zip_path: str):
    """
    Spawns a detached Windows batch script that:
    1. Waits for current AutoPrint-Outlook process to cleanly terminate.
    2. Extracts update files over the app directory, strictly excluding:
       - Downloads/
       - Printed_Archive/
       - printed_history.json
    3. Cleans up staging zip.
    4. Restarts the updated application.
    """
    updater_bat = os.path.join(app_dir, "_run_update.bat")
    
    # Script content: Safe in-place binary upgrade leaving all user data untouched
    bat_content = f"""@echo off
title Updating AutoPrint Outlook...
echo ==============================================================
echo           AutoPrint Outlook - Safe In-Place Updater
echo ==============================================================
echo.
echo 1. Waiting for application to close...
timeout /t 2 /nobreak >nul

echo 2. Extracting updated program files...
powershell -Command "Expand-Archive -LiteralPath '{update_zip_path}' -DestinationPath '{os.path.join(app_dir, '_update_staging', 'extracted')}' -Force"

echo 3. Replacing program binaries (Preserving all user data)...
powershell -Command "Get-ChildItem -Path '{os.path.join(app_dir, '_update_staging', 'extracted')}' -Recurse | Where-Object {{ -not ($_.FullName -match 'Downloads' -or $_.FullName -match 'Printed_Archive' -or $_.Name -match 'printed_history.json') }} | ForEach-Object {{ if (-not $_.PSIsContainer) {{ $rel = $_.FullName.Substring('{os.path.join(app_dir, '_update_staging', 'extracted')}'.Length + 1); $dest = Join-Path '{app_dir}' $rel; New-Item -ItemType Directory -Path (Split-Path $dest) -Force -ErrorAction SilentlyContinue | Out-Null; Copy-Item -Force $_.FullName $dest }} }}"

echo 4. Cleaning up temporary update payload...
rmdir /s /q "{os.path.join(app_dir, '_update_staging')}" >nul 2>&1

echo 5. Relaunching updated AutoPrint Outlook...
cd /d "{app_dir}"
if exist "AutoPrint-Outlook.exe" (
    start "" "AutoPrint-Outlook.exe"
) else (
    start "" python app.py
)

del "%~f0"
exit
"""
    with open(updater_bat, "w", encoding="utf-8") as f:
        f.write(bat_content)
        
    # Launch detached process
    subprocess.Popen(["cmd.exe", "/c", updater_bat], close_fds=True)
    # Terminate current app to release file lock
    sys.exit(0)
