@echo off
setlocal
chcp 65001 >nul
cd /d %~dp0

if not exist node_modules (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 goto :error
)

echo.
echo Запускаю Web UI на http://localhost:3000
echo Після запуску відкрийте браузер, якщо він не відкрився автоматично.
echo Для зупинки натисніть Ctrl+C у цьому вікні.
echo.
start "" "http://localhost:3000"
node server.js
if errorlevel 1 goto :error

exit /b 0

:error
echo.
echo Не вдалося запустити Web UI.
pause
exit /b 1
