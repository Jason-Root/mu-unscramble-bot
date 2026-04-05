@echo off
cd /d "%~dp0"

if exist "%~dp0Start MU Unscramble Bot.vbs" (
  start "" wscript.exe "%~dp0Start MU Unscramble Bot.vbs"
  exit /b 0
)

if not exist ".venv\Scripts\pythonw.exe" (
  echo Virtual environment not found at:
  echo   %cd%\.venv\Scripts\pythonw.exe
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process -Verb RunAs -FilePath '%cd%\.venv\Scripts\pythonw.exe' -WorkingDirectory '%cd%' -ArgumentList '-m mu_unscramble_bot'"
