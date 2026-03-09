@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

if not exist node_modules (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 goto :error
)

echo.
echo Starting Web UI at http://localhost:3000
echo If browser did not open automatically, open the URL manually.
echo Press Ctrl+C in this window to stop the server.
echo.
start "" "http://localhost:3000"
node server.js
if errorlevel 1 goto :error

exit /b 0

:error
echo.
echo Failed to start Web UI.
pause
exit /b 1
