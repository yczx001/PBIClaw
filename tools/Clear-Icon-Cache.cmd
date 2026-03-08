@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%Clear-Icon-Cache.ps1"

if not exist "%PS_SCRIPT%" (
  echo [ERROR] Script not found: "%PS_SCRIPT%"
  pause
  exit /b 1
)

echo Running icon cache cleanup (deep mode)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -Deep
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
  echo Script failed, exit code: %EXIT_CODE%
) else (
  echo Script completed.
)

pause
exit /b %EXIT_CODE%
