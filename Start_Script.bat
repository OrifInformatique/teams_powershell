@echo off

:: Check for admin privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin privileges...
    powershell.exe -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Run the script
powershell.exe -NoExit -ExecutionPolicy Bypass -File "%~dp0src\_main.ps1"