@echo off
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
  py -3 start_connected.py
) else (
  python start_connected.py
)
if errorlevel 1 (
  echo.
  echo Python 3 is needed to run the local Excel-connected dashboard.
  pause
)
