@echo off
cd /d %~dp0
if not exist node_modules (
  echo Installing dependencies...
  call npm install
)
set INPUT=Розділ_3_Практичний_аналіз_функціонування_іншомовної_лексики_в_корейських.docx
set OUTPUT=output.docx
node formatDocx.js "%INPUT%" "%OUTPUT%"
node verify.js "%OUTPUT%"
pause
