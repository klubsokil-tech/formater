@echo off
setlocal
chcp 65001 >nul
cd /d %~dp0

if not exist node_modules (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 goto :error
)

set "OUTPUT=output.docx"
set "DEFAULT_INPUT=Розділ_3_Практичний_аналіз_функціонування_іншомовної_лексики_в_корейських.docx"
set "INPUT="

if not "%~1"=="" (
  set "INPUT=%~1"
) else if exist "%DEFAULT_INPUT%" (
  set "INPUT=%DEFAULT_INPUT%"
) else if exist "input.docx" (
  set "INPUT=input.docx"
) else (
  for %%F in (*.docx) do (
    if /I not "%%~nxF"=="%OUTPUT%" (
      set "INPUT=%%~fF"
      goto :input_found
    )
  )
)

:input_found
if "%INPUT%"=="" (
  echo Помилка: не знайдено вхідний .docx файл.
  echo.
  echo Підказка: перетягніть файл на run-format.bat або покладіть input.docx у цю папку.
  goto :error
)

if not exist "%INPUT%" (
  echo Помилка: вхідний файл не існує:
  echo   %INPUT%
  goto :error
)

echo Вхідний файл: %INPUT%
echo Вихідний файл: %OUTPUT%

node formatDocx.js "%INPUT%" "%OUTPUT%"
if errorlevel 1 goto :error

node verify.js "%OUTPUT%"
if errorlevel 1 goto :error

echo.
echo Готово.
pause
exit /b 0

:error
echo.
echo Форматування завершено з помилкою.
pause
exit /b 1
