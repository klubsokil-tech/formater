@echo off
setlocal EnableExtensions
cd /d %~dp0

echo ==============================================
echo   Academic DOCX Formatter (Windows launcher)
echo ==============================================

echo [1/4] Перевірка Node.js...
where node >nul 2>nul
if errorlevel 1 goto :missing_node

echo [2/4] Перевірка npm...
where npm >nul 2>nul
if errorlevel 1 goto :missing_npm

if not exist node_modules (
  echo [3/4] Встановлення залежностей (npm install)...
  call npm install
  if errorlevel 1 goto :npm_install_failed
) else (
  echo [3/4] Залежності вже встановлені.
)

set "INPUT=Розділ_3_Практичний_аналіз_функціонування_іншомовної_лексики_в_корейських.docx"
set "OUTPUT=output.docx"

if not exist "%INPUT%" goto :missing_input

echo [4/4] Форматування документа...
node formatDocx.js "%INPUT%" "%OUTPUT%"
if errorlevel 1 goto :format_failed

echo Перевірка результату...
node verify.js "%OUTPUT%"
if errorlevel 1 goto :verify_failed

echo.
echo Готово! Створено файл: %OUTPUT%
goto :end

:missing_node
echo.
echo ПОМИЛКА: Node.js не знайдено у PATH.
echo Встановіть LTS-версію Node.js з https://nodejs.org/
echo Після встановлення закрийте і знову відкрийте це вікно, потім запустіть run-format.bat ще раз.
goto :end

:missing_npm
echo.
echo ПОМИЛКА: npm не знайдено у PATH.
echo Перевстановіть Node.js LTS з https://nodejs.org/ і переконайтесь, що npm встановлено.
goto :end

:npm_install_failed
echo.
echo ПОМИЛКА: не вдалося виконати npm install.
echo Перевірте підключення до інтернету та права доступу, потім повторіть спробу.
goto :end

:missing_input
echo.
echo ПОМИЛКА: вхідний файл не знайдено:
echo   %INPUT%
echo Помістіть файл у папку проєкту і запустіть run-format.bat знову.
goto :end

:format_failed
echo.
echo ПОМИЛКА: форматування завершилось з помилкою.
goto :end

:verify_failed
echo.
echo УВАГА: форматування завершено, але verify.js повідомив про невідповідності.
echo Перевірте повідомлення вище.
goto :end

:end
echo.
pause
endlocal
