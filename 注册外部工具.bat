@echo off
net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\register-tool.ps1"
pause
