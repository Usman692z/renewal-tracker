@echo off
cd /d "%~dp0"
echo === Step 1/2: Fresh pull from SharePoint (CSU Daily Report pipeline) ===
pushd "D:\CSU Daily Report"
py fetch_data.py && py build_dashboard.py
popd
echo.
echo === Step 2/2: Upload requests to Renewal Tracker (Firestore) ===
node "scripts\sync-requests.js"
echo.
pause
