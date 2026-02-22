@echo off
setlocal

REM Run from the repository directory where this script lives
cd /d "%~dp0"

REM Prefer venv Python if present, otherwise fall back to py/python
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" app.py
) else (
  where py >nul 2>nul
  if %errorlevel%==0 (
    py -3 app.py
  ) else (
    python app.py
  )
)

endlocal
