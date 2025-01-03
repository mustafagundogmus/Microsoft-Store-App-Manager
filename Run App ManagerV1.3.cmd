@echo off
:: Set the current directory as dp0
set "dp0=%~dp0"

:: Check if the script is run with administrator privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Administrator privileges are required. Restarting with administrator rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Run settings.ps1 file in hidden mode
echo Running settings.ps1 with administrator privileges...
powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File "%dp0%Microsoft Store App Manager.ps1"
if %errorlevel% NEQ 0 (
    echo An error occurred while running settings.ps1.
    pause
    exit /b
)

echo Process completed successfully.
exit /b
