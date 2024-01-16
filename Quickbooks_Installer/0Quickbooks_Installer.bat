@echo off
setlocal

:: Check if running with administrative privileges
>nul 2>&1 net session
if %errorLevel% neq 0 (
 echo This script requires administrative privileges. Please run as administrator.
 pause
 exit /b 1
)

:: Change the current directory to the batch file's directory
cd %~dp0

:: Run the PowerShell script in an elevated session
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0QB_install.ps1""' -Verb RunAs}"

:: The following line is needed to keep the console window open
pause

endlocal