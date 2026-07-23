@echo off
title Renewal Tracker - Auto Sync
cd /d "%~dp0"
echo ============================================================
echo  Auto Sync: pushes app edits to SharePoint, then pulls fresh
echo  SharePoint data back into the app, every 5 minutes.
echo  Keep this window open. Ctrl+C to stop.
echo ============================================================
:loop
echo.
echo [%date% %time%] Pushing app edits to SharePoint...
node "scripts\push-to-sharepoint.js" --apply
echo [%date% %time%] Pulling fresh data from SharePoint...
pushd "D:\CSU Daily Report"
py fetch_data.py && py build_dashboard.py
popd
echo [%date% %time%] Uploading to Renewal Tracker...
node "scripts\sync-requests.js"
echo [%date% %time%] Done. Next sync in 5 minutes...
timeout /t 300 /nobreak >nul
goto loop
