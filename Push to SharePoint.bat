@echo off
cd /d "%~dp0"
echo === Dry run first (shows what WOULD be written, writes nothing) ===
node "scripts\push-to-sharepoint.js"
echo.
set /p CONFIRM="Type YES to actually write these changes to SharePoint: "
if /I "%CONFIRM%"=="YES" (
  node "scripts\push-to-sharepoint.js" --apply
) else (
  echo Cancelled, nothing was written.
)
echo.
pause
