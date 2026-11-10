@echo off
chcp 65001 > nul
rem Drag a month-end P&L Excel file onto this file.
if "%~1"=="" (
  echo [!] Drag the month-end Excel file onto this .bat file.
  pause
  exit /b 1
)
python "%~dp0refresh.py" "%~1"
echo.
pause
